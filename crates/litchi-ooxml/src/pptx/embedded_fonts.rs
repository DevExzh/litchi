//! Typed PresentationML embedded-font references and inert OPC resources.

use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const FONT_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
const STRICT_FONT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/font";
const PRESENTATION_CT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const FONT_DATA_CT: &str = "application/x-fontdata";
const FONT_TTF_CT: &str = "application/x-font-ttf";
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_DEPTH: usize = 256;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_FONTS: usize = 4096;
const MAX_FONT_BYTES: usize = 64 * 1024 * 1024;
const MAX_TOTAL_FONT_BYTES: usize = 256 * 1024 * 1024;

/// Validated OpenType `OS/2.fsType` embedding metadata.
///
/// This value is supplied by callers or a separate font-metadata reader. This
/// module never searches, parses, loads, or executes a font program.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EmbeddedFontLicensing {
    pub fs_type: u16,
    pub restricted_license: bool,
    pub preview_and_print: bool,
    pub editable: bool,
    pub no_subsetting: bool,
    pub bitmap_only: bool,
}

impl EmbeddedFontLicensing {
    /// Validate the defined embedding bits and reject contradictory modes.
    pub fn from_fs_type(fs_type: u16) -> Result<Self> {
        const DEFINED: u16 = 0x0002 | 0x0004 | 0x0008 | 0x0100 | 0x0200;
        if fs_type & !DEFINED != 0 {
            return Err(invalid(format!(
                "font fsType contains reserved bits 0x{:04X}",
                fs_type & !DEFINED
            )));
        }
        let modes = [0x0002, 0x0004, 0x0008]
            .into_iter()
            .filter(|bit| fs_type & *bit != 0)
            .count();
        if modes > 1 {
            return Err(invalid(
                "font fsType has contradictory restricted, preview/print, and editable modes",
            ));
        }
        Ok(Self {
            fs_type,
            restricted_license: fs_type & 0x0002 != 0,
            preview_and_print: fs_type & 0x0004 != 0,
            editable: fs_type & 0x0008 != 0,
            no_subsetting: fs_type & 0x0100 != 0,
            bitmap_only: fs_type & 0x0200 != 0,
        })
    }

    /// Bit zero through three are clear for installable embedding.
    pub fn installable(self) -> bool {
        self.fs_type & 0x000E == 0
    }
}

/// Apply the reversible OOXML GUID XOR transformation to the first 32 bytes.
///
/// The font program remains inert bytes. The operation rejects short input and
/// malformed keys rather than attempting to inspect the font.
pub fn obfuscate_embedded_font_data(data: &mut [u8], font_key: &str) -> Result<()> {
    if data.len() < 32 {
        return Err(invalid("OOXML font obfuscation requires at least 32 bytes"));
    }
    let key = parse_font_key(font_key)?;
    for index in 0..32 {
        data[index] ^= key[15 - (index % 16)];
    }
    Ok(())
}

/// Reverse [`obfuscate_embedded_font_data`]. XOR makes both operations identical.
pub fn deobfuscate_embedded_font_data(data: &mut [u8], font_key: &str) -> Result<()> {
    obfuscate_embedded_font_data(data, font_key)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EmbeddedFontConformance {
    Transitional,
    Strict,
}

impl EmbeddedFontConformance {
    fn pml(self) -> &'static str {
        match self {
            Self::Transitional => PML,
            Self::Strict => STRICT_PML,
        }
    }
    fn rel_ns(self) -> &'static str {
        match self {
            Self::Transitional => REL_NS,
            Self::Strict => STRICT_REL_NS,
        }
    }
    fn font_rel(self) -> &'static str {
        match self {
            Self::Transitional => FONT_REL,
            Self::Strict => STRICT_FONT_REL,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum EmbeddedFontStyle {
    Regular,
    Bold,
    Italic,
    BoldItalic,
}

impl EmbeddedFontStyle {
    fn element(self) -> &'static str {
        match self {
            Self::Regular => "regular",
            Self::Bold => "bold",
            Self::Italic => "italic",
            Self::BoldItalic => "boldItalic",
        }
    }
    fn rank(self) -> u8 {
        match self {
            Self::Regular => 0,
            Self::Bold => 1,
            Self::Italic => 2,
            Self::BoldItalic => 3,
        }
    }
    fn parse(value: &str) -> Option<Self> {
        match value {
            "regular" => Some(Self::Regular),
            "bold" => Some(Self::Bold),
            "italic" => Some(Self::Italic),
            "boldItalic" => Some(Self::BoldItalic),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedFontResource {
    pub part_name: String,
    pub content_type: String,
    /// The font program is deliberately retained as inert bytes.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedFontFace {
    pub style: EmbeddedFontStyle,
    pub relationship_id: String,
    /// Present after package loading and required for package storage.
    pub resource: Option<EmbeddedFontResource>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedFont {
    pub typeface: String,
    pub panose: Option<[u8; 10]>,
    pub pitch_family: Option<u8>,
    pub charset: Option<u8>,
    pub faces: Vec<EmbeddedFontFace>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PresentationEmbeddedFonts {
    pub fonts: Vec<EmbeddedFont>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Presentation,
    List,
    Font(usize),
    Leaf,
    Other,
}

struct ParsedPresentation {
    conformance: EmbeddedFontConformance,
    value: Option<PresentationEmbeddedFonts>,
}

/// Parses the optional `p:embeddedFontLst` from a complete presentation part.
pub fn parse_embedded_fonts(xml: &[u8]) -> Result<Option<PresentationEmbeddedFonts>> {
    Ok(parse_presentation(xml)?.value)
}

fn parse_presentation(xml: &[u8]) -> Result<ParsedPresentation> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("presentation XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_XML_BYTES {
        return Err(limit("MCE-processed presentation XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Context> = Vec::new();
    let mut fonts = Vec::new();
    let mut saw_root = false;
    let mut saw_list = false;
    let mut conformance = None;
    let mut nodes = 0usize;
    let mut string_bytes = 0usize;
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        let empty_event = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = empty_event;
                nodes = nodes.checked_add(1).ok_or_else(|| limit("XML nodes"))?;
                if nodes > MAX_NODES {
                    return Err(limit("XML nodes"));
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
                let namespace =
                    resolved_namespace(reader.resolver().resolve_element(element.name()).0)?;
                let local = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                let parent = stack.last().copied();
                let context = if parent.is_none() {
                    if saw_root || local != "presentation" {
                        return Err(invalid("expected one PresentationML presentation root"));
                    }
                    let c = match namespace.as_str() {
                        PML => EmbeddedFontConformance::Transitional,
                        STRICT_PML => EmbeddedFontConformance::Strict,
                        _ => return Err(invalid("presentation root has an unsupported namespace")),
                    };
                    saw_root = true;
                    conformance = Some(c);
                    Context::Presentation
                } else if parent == Some(Context::Presentation)
                    && namespace == conformance.unwrap().pml()
                    && local == "embeddedFontLst"
                {
                    if saw_list {
                        return Err(invalid(
                            "presentation has multiple embeddedFontLst elements",
                        ));
                    }
                    reject_unqualified_attributes(
                        &reader,
                        element,
                        reader.decoder(),
                        &[],
                        &mut string_bytes,
                    )?;
                    saw_list = true;
                    Context::List
                } else if parent == Some(Context::List) {
                    if namespace != conformance.unwrap().pml() || local != "embeddedFont" {
                        return Err(invalid("embeddedFontLst contains a non-embeddedFont child"));
                    }
                    reject_unqualified_attributes(
                        &reader,
                        element,
                        reader.decoder(),
                        &[],
                        &mut string_bytes,
                    )?;
                    if fonts.len() >= MAX_FONTS {
                        return Err(limit("embedded fonts"));
                    }
                    fonts.push(EmbeddedFont {
                        typeface: String::new(),
                        panose: None,
                        pitch_family: None,
                        charset: None,
                        faces: Vec::new(),
                    });
                    Context::Font(fonts.len() - 1)
                } else if let Some(Context::Font(index)) = parent {
                    if namespace != conformance.unwrap().pml() {
                        return Err(invalid("embeddedFont contains a foreign child"));
                    }
                    if local == "font" {
                        if !fonts[index].typeface.is_empty() {
                            return Err(invalid("embeddedFont has multiple font descriptors"));
                        }
                        parse_descriptor(
                            &reader,
                            element,
                            reader.decoder(),
                            &mut fonts[index],
                            &mut string_bytes,
                        )?;
                        Context::Leaf
                    } else if let Some(style) = EmbeddedFontStyle::parse(&local) {
                        let relationship_id =
                            parse_face(&reader, element, reader.decoder(), &mut string_bytes)?;
                        if fonts[index].faces.iter().any(|face| face.style == style) {
                            return Err(invalid(format!(
                                "duplicate embedded-font style '{local}'"
                            )));
                        }
                        if fonts[index]
                            .faces
                            .last()
                            .is_some_and(|face| face.style.rank() >= style.rank())
                        {
                            return Err(invalid("embedded-font styles are out of schema order"));
                        }
                        fonts[index].faces.push(EmbeddedFontFace {
                            style,
                            relationship_id,
                            resource: None,
                        });
                        Context::Leaf
                    } else {
                        return Err(invalid(format!("unexpected embeddedFont child '{local}'")));
                    }
                } else if matches!(parent, Some(Context::List | Context::Leaf)) {
                    return Err(invalid("embedded-font leaf element contains child content"));
                } else {
                    Context::Other
                };
                stack.push(context);
                if empty {
                    let ended = stack.pop().unwrap();
                    if let Context::Font(index) = ended {
                        finish_font(&fonts[index])?;
                    }
                }
            },
            Event::End(_) => {
                let ended = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                if let Context::Font(index) = ended {
                    finish_font(&fonts[index])?;
                }
            },
            Event::Text(text) => {
                if matches!(
                    stack.last(),
                    Some(Context::List | Context::Font(_) | Context::Leaf)
                ) {
                    let value = text.decode().map_err(xml_error)?;
                    if !value.trim().is_empty() {
                        return Err(invalid("embedded-font markup contains text"));
                    }
                } else if stack.is_empty() && !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("text occurs outside the presentation root"));
                }
            },
            Event::CData(_) => {
                if matches!(
                    stack.last(),
                    Some(Context::List | Context::Font(_) | Context::Leaf)
                ) {
                    return Err(invalid("CDATA is rejected in embedded-font markup"));
                }
            },
            Event::GeneralRef(_) => {
                if matches!(
                    stack.last(),
                    Some(Context::List | Context::Font(_) | Context::Leaf)
                ) {
                    return Err(invalid(
                        "entity references are rejected in embedded-font markup",
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation XML"));
    }
    let conformance = conformance.ok_or_else(|| invalid("missing presentation root"))?;
    let value = saw_list.then_some(PresentationEmbeddedFonts { fonts });
    if let Some(value) = &value {
        validate_value(value, false)?;
    }
    Ok(ParsedPresentation { conformance, value })
}

fn parse_descriptor(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    font: &mut EmbeddedFont,
    strings: &mut usize,
) -> Result<()> {
    let attrs = collect_unqualified_attributes(
        reader,
        element,
        decoder,
        &["typeface", "panose", "pitchFamily", "charset"],
        strings,
    )?;
    font.typeface = attrs
        .get("typeface")
        .cloned()
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid("font descriptor is missing typeface"))?;
    font.panose = attrs
        .get("panose")
        .map(|value| parse_panose(value))
        .transpose()?;
    font.pitch_family = attrs
        .get("pitchFamily")
        .map(|value| parse_u8(value, "pitchFamily"))
        .transpose()?;
    font.charset = attrs
        .get("charset")
        .map(|value| parse_u8(value, "charset"))
        .transpose()?;
    Ok(())
}

fn parse_face(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    strings: &mut usize,
) -> Result<String> {
    let mut id = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let qualified = attribute.key.as_ref();
        if qualified == b"xmlns" || qualified.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if local.as_ref() != b"id"
            || !matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == REL_NS.as_bytes() || value == STRICT_REL_NS.as_bytes())
        {
            return Err(invalid("embedded-font face has an unexpected attribute"));
        }
        if id.is_some() {
            return Err(invalid("embedded-font face has duplicate relationship IDs"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_string_bytes(strings, value.len())?;
        id = Some(value);
    }
    let id = id
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid("embedded-font face is missing r:id"))?;
    validate_relationship_id(&id)?;
    Ok(id)
}

fn collect_unqualified_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    allowed: &[&str],
    strings: &mut usize,
) -> Result<HashMap<String, String>> {
    let mut result = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let qualified = attribute.key.as_ref();
        if qualified == b"xmlns" || qualified.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let local = std::str::from_utf8(local.as_ref()).map_err(xml_error)?;
        if namespace != ResolveResult::Unbound || !allowed.contains(&local) {
            return Err(invalid(format!("unexpected attribute '{local}'")));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_string_bytes(strings, local.len() + value.len())?;
        if result.insert(local.to_owned(), value).is_some() {
            return Err(invalid(format!("duplicate attribute '{local}'")));
        }
    }
    Ok(result)
}

fn reject_unqualified_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    allowed: &[&str],
    strings: &mut usize,
) -> Result<()> {
    collect_unqualified_attributes(reader, element, decoder, allowed, strings).map(|_| ())
}

/// Deterministically serializes a self-contained `p:embeddedFontLst` fragment.
pub fn write_embedded_font_list(
    value: &PresentationEmbeddedFonts,
    conformance: EmbeddedFontConformance,
) -> Result<Vec<u8>> {
    validate_value(value, false)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<p:embeddedFontLst xmlns:p=\"");
    escape(&mut output, conformance.pml());
    output.extend_from_slice(b"\" xmlns:r=\"");
    escape(&mut output, conformance.rel_ns());
    if value.fonts.is_empty() {
        output.extend_from_slice(b"\"/>");
        return Ok(output);
    }
    output.extend_from_slice(b"\">");
    for font in &value.fonts {
        output.extend_from_slice(b"<p:embeddedFont><p:font typeface=\"");
        escape(&mut output, &font.typeface);
        output.push(b'\"');
        if let Some(panose) = font.panose {
            attribute(&mut output, "panose", &hex_panose(&panose));
        }
        if let Some(value) = font.pitch_family {
            attribute(&mut output, "pitchFamily", &value.to_string());
        }
        if let Some(value) = font.charset {
            attribute(&mut output, "charset", &value.to_string());
        }
        output.extend_from_slice(b"/>");
        for face in &font.faces {
            output.extend_from_slice(b"<p:");
            output.extend_from_slice(face.style.element().as_bytes());
            output.extend_from_slice(b" r:id=\"");
            escape(&mut output, &face.relationship_id);
            output.extend_from_slice(b"\"/>");
        }
        output.extend_from_slice(b"</p:embeddedFont>");
    }
    output.extend_from_slice(b"</p:embeddedFontLst>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized embedded-font XML bytes"));
    }
    Ok(output)
}

/// Loads embedded-font metadata and validates every referenced inert font part.
pub fn load_embedded_fonts(package: &OpcPackage) -> Result<Option<PresentationEmbeddedFonts>> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation)?;
    let presentation_name = presentation.partname().to_string();
    let parsed = parse_presentation(presentation.blob())?;
    validate_font_relationship_sources(package, &presentation_name)?;
    let Some(mut value) = parsed.value else {
        reject_orphan_font_parts(package, &HashSet::new())?;
        return Ok(None);
    };
    let mut targets = HashSet::new();
    let mut references = HashSet::new();
    let mut resources = HashMap::<String, EmbeddedFontResource>::new();
    let mut total_bytes = 0usize;
    for font in &mut value.fonts {
        for face in &mut font.faces {
            if !references.insert(face.relationship_id.clone()) {
                return Err(invalid(format!(
                    "duplicate embedded-font relationship reference '{}'",
                    face.relationship_id
                )));
            }
            let relationship = presentation
                .rels()
                .get(&face.relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "missing embedded-font relationship '{}'",
                        face.relationship_id
                    ))
                })?;
            if !is_font_relationship(relationship.reltype()) {
                return Err(invalid(format!(
                    "relationship '{}' is not a font relationship",
                    face.relationship_id
                )));
            }
            if relationship.is_external() {
                return Err(invalid("embedded-font relationship must be internal"));
            }
            let target = relationship.target_partname()?;
            let target_name = target.to_string();
            targets.insert(target_name.clone());
            if let Some(resource) = resources.get(&target_name) {
                face.resource = Some(resource.clone());
                continue;
            }
            if !target.as_str().starts_with("/ppt/fonts/") {
                return Err(invalid(format!(
                    "font part '{target}' is outside /ppt/fonts"
                )));
            }
            let part = package.get_part(&target)?;
            if !is_font_content_type(part.content_type()) {
                return Err(invalid(format!(
                    "font part '{target}' has invalid content type '{}'",
                    part.content_type()
                )));
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "font part '{target}' has outbound relationships"
                )));
            }
            if part.blob().len() > MAX_FONT_BYTES {
                return Err(limit("individual font bytes"));
            }
            total_bytes = total_bytes
                .checked_add(part.blob().len())
                .ok_or_else(|| limit("total font bytes"))?;
            if total_bytes > MAX_TOTAL_FONT_BYTES {
                return Err(limit("total font bytes"));
            }
            let resource = EmbeddedFontResource {
                part_name: target_name.clone(),
                content_type: part.content_type().to_owned(),
                data: part.blob().to_vec(),
            };
            resources.insert(target_name, resource.clone());
            face.resource = Some(resource);
        }
    }
    validate_inbound_font_graph(
        package,
        &presentation_name,
        presentation,
        &references,
        &targets,
    )?;
    reject_orphan_font_parts(package, &targets)?;
    Ok(Some(value))
}

/// Atomically stores the complete embedded-font graph.
///
/// Existing font relationships are replaced. Font parts still referenced by
/// another relationship are retained, and unrelated presentation XML is copied
/// byte-for-byte.
pub fn store_embedded_fonts(
    package: &mut OpcPackage,
    value: &PresentationEmbeddedFonts,
    conformance: EmbeddedFontConformance,
) -> Result<()> {
    validate_value(value, true)?;
    let old = load_embedded_fonts(package)?.unwrap_or_default();
    let presentation = package.main_document_part()?;
    let presentation_name = presentation.partname().clone();
    let parsed = parse_presentation(presentation.blob())?;
    if parsed.conformance != conformance {
        return Err(invalid(
            "requested conformance does not match the presentation namespace",
        ));
    }
    let fragment = if value.fonts.is_empty() {
        Vec::new()
    } else {
        write_embedded_font_list(value, conformance)?
    };
    let updated_xml = patch_font_list(presentation.blob(), &fragment, conformance)?;
    let staged = parse_presentation(&updated_xml)?;
    if staged.conformance != conformance || staged.value.unwrap_or_default() != metadata_only(value)
    {
        return Err(invalid("staged embedded-font XML did not round-trip"));
    }
    let old_relationship_ids = old
        .fonts
        .iter()
        .flat_map(|font| font.faces.iter().map(|face| face.relationship_id.clone()))
        .collect::<HashSet<_>>();
    let old_part_names = old
        .fonts
        .iter()
        .flat_map(|font| font.faces.iter())
        .filter_map(|face| {
            face.resource
                .as_ref()
                .map(|resource| resource.part_name.clone())
        })
        .collect::<HashSet<_>>();
    let mut relationship_ids = HashSet::new();
    let mut resources = HashMap::<String, (String, Vec<u8>)>::new();
    let mut relationships = Vec::new();
    for font in &value.fonts {
        for face in &font.faces {
            if !relationship_ids.insert(face.relationship_id.clone()) {
                return Err(invalid(format!(
                    "duplicate relationship ID '{}'",
                    face.relationship_id
                )));
            }
            if presentation.rels().get(&face.relationship_id).is_some()
                && !old_relationship_ids.contains(&face.relationship_id)
            {
                return Err(invalid(format!(
                    "relationship ID '{}' already exists",
                    face.relationship_id
                )));
            }
            let resource = face
                .resource
                .as_ref()
                .ok_or_else(|| invalid("embedded-font resource is required for package storage"))?;
            let uri = PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?;
            if !uri.as_str().starts_with("/ppt/fonts/") {
                return Err(invalid(format!("font part '{uri}' is outside /ppt/fonts")));
            }
            if let Some((content_type, data)) = resources.get(uri.as_str()) {
                if content_type != &resource.content_type || data != &resource.data {
                    return Err(invalid(format!(
                        "shared font part '{uri}' has conflicting resources"
                    )));
                }
            } else {
                resources.insert(
                    uri.to_string(),
                    (resource.content_type.clone(), resource.data.clone()),
                );
            }
            relationships.push((uri, face.relationship_id.clone()));
        }
    }

    for (part_name, (content_type, data)) in &resources {
        let uri = PackURI::new(part_name).map_err(OoxmlError::InvalidUri)?;
        if let Ok(part) = package.get_part(&uri) {
            if part.content_type() != content_type {
                return Err(invalid(format!("font part '{uri}' content type collision")));
            }
            if !part.rels().is_empty() {
                return Err(invalid(format!(
                    "font part '{uri}' has outbound relationships"
                )));
            }
            if part.blob() != data && !old_part_names.contains(part_name) {
                return Err(invalid(format!("font part '{uri}' data collision")));
            }
            if part.blob() != data
                && has_inbound_outside_relationships(
                    package,
                    &uri,
                    &presentation_name,
                    &old_relationship_ids,
                )?
            {
                return Err(invalid(format!(
                    "shared font part '{uri}' cannot be overwritten"
                )));
            }
        }
    }

    package.unsign();
    let existing_font_relationships = package
        .get_part(&presentation_name)?
        .rels()
        .iter()
        .filter(|relationship| is_font_relationship(relationship.reltype()))
        .map(|relationship| relationship.r_id().to_owned())
        .collect::<Vec<_>>();
    for relationship_id in existing_font_relationships {
        package
            .get_part_mut(&presentation_name)?
            .rels_mut()
            .remove(&relationship_id);
    }
    for (uri, relationship_id) in &relationships {
        package
            .get_part_mut(&presentation_name)?
            .rels_mut()
            .add_relationship(
                conformance.font_rel().into(),
                uri.relative_ref(presentation_name.base_uri()),
                relationship_id.clone(),
                false,
            );
    }
    for (part_name, (content_type, data)) in resources {
        let uri = PackURI::new(&part_name).map_err(OoxmlError::InvalidUri)?;
        if let Ok(part) = package.get_part_mut(&uri) {
            part.set_blob(data);
        } else {
            package.add_part(Box::new(BlobPart::new(uri, content_type, data)));
        }
    }
    package
        .get_part_mut(&presentation_name)?
        .set_blob(updated_xml);
    let retained = relationships
        .iter()
        .map(|(uri, _)| uri.to_string())
        .collect::<HashSet<_>>();
    for old_part in old_part_names {
        if !retained.contains(&old_part) {
            let uri = PackURI::new(&old_part).map_err(OoxmlError::InvalidUri)?;
            if !part_is_referenced(package, &uri)? {
                package.remove_part(&uri);
            }
        }
    }
    Ok(())
}

/// Find an embedded typeface using PowerPoint's case-insensitive identity.
pub fn find_embedded_font(package: &OpcPackage, typeface: &str) -> Result<Option<EmbeddedFont>> {
    Ok(load_embedded_fonts(package)?.and_then(|value| {
        value
            .fonts
            .into_iter()
            .find(|font| font.typeface.eq_ignore_ascii_case(typeface))
    }))
}

/// Add a typeface, allocating blank relationship IDs and part names safely.
pub fn add_embedded_font(
    package: &mut OpcPackage,
    mut font: EmbeddedFont,
    conformance: EmbeddedFontConformance,
) -> Result<()> {
    let mut value = load_embedded_fonts(package)?.unwrap_or_default();
    if value
        .fonts
        .iter()
        .any(|item| item.typeface.eq_ignore_ascii_case(&font.typeface))
    {
        return Err(invalid(format!(
            "embedded font typeface '{}' already exists",
            font.typeface
        )));
    }
    allocate_font_identifiers(package, &mut font, &value)?;
    value.fonts.push(font);
    store_embedded_fonts(package, &value, conformance)
}

/// Update a typeface selected by its current name.
pub fn update_embedded_font(
    package: &mut OpcPackage,
    typeface: &str,
    mut replacement: EmbeddedFont,
    conformance: EmbeddedFontConformance,
) -> Result<()> {
    let mut value = load_embedded_fonts(package)?
        .ok_or_else(|| invalid("presentation has no embedded fonts"))?;
    let offset = value
        .fonts
        .iter()
        .position(|font| font.typeface.eq_ignore_ascii_case(typeface))
        .ok_or_else(|| invalid(format!("embedded font '{typeface}' was not found")))?;
    value.fonts.remove(offset);
    allocate_font_identifiers(package, &mut replacement, &value)?;
    value.fonts.insert(offset, replacement);
    store_embedded_fonts(package, &value, conformance)
}

pub fn replace_embedded_font(
    package: &mut OpcPackage,
    typeface: &str,
    replacement: EmbeddedFont,
    conformance: EmbeddedFontConformance,
) -> Result<()> {
    update_embedded_font(package, typeface, replacement, conformance)
}

pub fn remove_embedded_font(
    package: &mut OpcPackage,
    typeface: &str,
    conformance: EmbeddedFontConformance,
) -> Result<bool> {
    let Some(mut value) = load_embedded_fonts(package)? else {
        return Ok(false);
    };
    let Some(offset) = value
        .fonts
        .iter()
        .position(|font| font.typeface.eq_ignore_ascii_case(typeface))
    else {
        return Ok(false);
    };
    value.fonts.remove(offset);
    store_embedded_fonts(package, &value, conformance)?;
    Ok(true)
}

pub fn reorder_embedded_fonts(
    package: &mut OpcPackage,
    ordered_typefaces: &[String],
    conformance: EmbeddedFontConformance,
) -> Result<()> {
    let mut value = load_embedded_fonts(package)?
        .ok_or_else(|| invalid("presentation has no embedded fonts"))?;
    let expected = value
        .fonts
        .iter()
        .map(|font| font.typeface.to_lowercase())
        .collect::<HashSet<_>>();
    let actual = ordered_typefaces
        .iter()
        .map(|typeface| typeface.to_lowercase())
        .collect::<HashSet<_>>();
    if expected != actual || ordered_typefaces.len() != value.fonts.len() {
        return Err(invalid(
            "embedded-font reorder is not a typeface permutation",
        ));
    }
    value.fonts = ordered_typefaces
        .iter()
        .map(|typeface| {
            value
                .fonts
                .iter()
                .find(|font| font.typeface.eq_ignore_ascii_case(typeface))
                .expect("permutation was validated")
                .clone()
        })
        .collect();
    store_embedded_fonts(package, &value, conformance)
}

fn allocate_font_identifiers(
    package: &OpcPackage,
    font: &mut EmbeddedFont,
    existing: &PresentationEmbeddedFonts,
) -> Result<()> {
    let presentation = package.main_document_part()?;
    let mut relationship_ids = presentation
        .rels()
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .chain(
            existing
                .fonts
                .iter()
                .flat_map(|font| font.faces.iter().map(|face| face.relationship_id.clone())),
        )
        .collect::<HashSet<_>>();
    let mut part_names = package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .chain(existing.fonts.iter().flat_map(|font| {
            font.faces.iter().filter_map(|face| {
                face.resource
                    .as_ref()
                    .map(|resource| resource.part_name.clone())
            })
        }))
        .collect::<HashSet<_>>();
    for face in &mut font.faces {
        if face.relationship_id.is_empty() {
            face.relationship_id = next_font_relationship_id(&relationship_ids)?;
        }
        relationship_ids.insert(face.relationship_id.clone());
        let resource = face
            .resource
            .as_mut()
            .ok_or_else(|| invalid("embedded-font resource is required"))?;
        if resource.part_name.is_empty() {
            resource.part_name = next_font_part_name(&part_names)?;
        }
        part_names.insert(resource.part_name.clone());
        if resource.content_type.is_empty() {
            resource.content_type = FONT_DATA_CT.into();
        }
    }
    Ok(())
}

fn next_font_relationship_id(used: &HashSet<String>) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("rIdFont{index}");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(limit("relationship IDs"))
}

fn next_font_part_name(used: &HashSet<String>) -> Result<String> {
    for index in 1..=u32::MAX {
        let candidate = format!("/ppt/fonts/font{index}.fntdata");
        if !used.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(limit("part names"))
}

fn patch_font_list(
    xml: &[u8],
    fragment: &[u8],
    conformance: EmbeddedFontConformance,
) -> Result<Vec<u8>> {
    let ranges = font_list_ranges(xml)?;
    if ranges.is_empty() {
        if fragment.is_empty() {
            return Ok(xml.to_vec());
        }
        return insert_font_list(xml, fragment, conformance);
    }
    let mut output = xml.to_vec();
    for (start, end) in ranges.into_iter().rev() {
        output.splice(start..end, fragment.iter().copied());
    }
    if output.len() > MAX_XML_BYTES {
        return Err(limit("updated presentation XML bytes"));
    }
    Ok(output)
}

fn font_list_ranges(xml: &[u8]) -> Result<Vec<(usize, usize)>> {
    let mut reader = NsReader::from_reader(xml);
    let mut starts = Vec::<Option<usize>>::new();
    let mut ranges = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("presentation XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let target = element.local_name().as_ref() == b"embeddedFontLst"
                    && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PML.as_bytes() || value == STRICT_PML.as_bytes());
                starts.push(target.then_some(start));
                if starts.len() > MAX_DEPTH {
                    return Err(limit("presentation XML depth"));
                }
            },
            Event::Empty(element)
                if element.local_name().as_ref() == b"embeddedFontLst"
                    && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PML.as_bytes() || value == STRICT_PML.as_bytes()) =>
            {
                ranges.push((
                    start,
                    usize::try_from(reader.buffer_position())
                        .map_err(|_| invalid("presentation XML offset overflow"))?,
                ));
            },
            Event::End(_) => {
                let target = starts
                    .pop()
                    .ok_or_else(|| invalid("unexpected presentation closing element"))?;
                if let Some(start) = target {
                    ranges.push((
                        start,
                        usize::try_from(reader.buffer_position())
                            .map_err(|_| invalid("presentation XML offset overflow"))?,
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !starts.is_empty() {
        return Err(invalid("unterminated presentation XML"));
    }
    ranges.sort_unstable_by_key(|range| range.0);
    for pair in ranges.windows(2) {
        if pair[0].1 > pair[1].0 {
            return Err(invalid("overlapping embedded-font XML ranges"));
        }
    }
    Ok(ranges)
}

fn insert_font_list(
    xml: &[u8],
    fragment: &[u8],
    conformance: EmbeddedFontConformance,
) -> Result<Vec<u8>> {
    let later = [
        b"custShowLst".as_slice(),
        b"photoAlbum",
        b"custDataLst",
        b"kinsoku",
        b"defaultTextStyle",
        b"modifyVerifier",
        b"extLst",
    ];
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut position = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("presentation XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let is_pml = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == conformance.pml().as_bytes());
                if depth == 0 {
                    if !is_pml || element.local_name().as_ref() != b"presentation" {
                        return Err(invalid(
                            "presentation root does not match requested conformance",
                        ));
                    }
                    root_seen = true;
                } else if depth == 1 && is_pml && later.contains(&element.local_name().as_ref()) {
                    position.get_or_insert(start);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("presentation XML depth"))?;
            },
            Event::Empty(element) => {
                let is_pml = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == conformance.pml().as_bytes());
                if depth == 1 && is_pml && later.contains(&element.local_name().as_ref()) {
                    position.get_or_insert(start);
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected presentation closing element"));
                }
                if depth == 1 && element.local_name().as_ref() == b"presentation" {
                    position.get_or_insert(start);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || depth != 0 {
        return Err(invalid("invalid presentation XML"));
    }
    let position = position.ok_or_else(|| invalid("missing presentation closing element"))?;
    let length = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| limit("updated presentation XML bytes"))?;
    if length > MAX_XML_BYTES {
        return Err(limit("updated presentation XML bytes"));
    }
    let mut output = Vec::with_capacity(length);
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

fn validate_value(value: &PresentationEmbeddedFonts, require_resources: bool) -> Result<()> {
    if value.fonts.len() > MAX_FONTS {
        return Err(limit("embedded fonts"));
    }
    let mut typefaces = HashSet::new();
    let mut total = 0usize;
    for font in &value.fonts {
        if font.typeface.is_empty() {
            return Err(invalid("embedded font typeface cannot be empty"));
        }
        bounded_string(&font.typeface)?;
        if !typefaces.insert(font.typeface.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate embedded font typeface '{}'",
                font.typeface
            )));
        }
        finish_font(font)?;
        for face in &font.faces {
            validate_relationship_id(&face.relationship_id)?;
            if require_resources && face.resource.is_none() {
                return Err(invalid(
                    "embedded-font resource is required for package storage",
                ));
            }
            if let Some(resource) = &face.resource {
                let uri = PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?;
                if !uri.as_str().starts_with("/ppt/fonts/") {
                    return Err(invalid(format!("font part '{uri}' is outside /ppt/fonts")));
                }
                if !is_font_content_type(&resource.content_type) {
                    return Err(invalid(format!(
                        "invalid embedded-font content type '{}'",
                        resource.content_type
                    )));
                }
                if resource.data.len() > MAX_FONT_BYTES {
                    return Err(limit("individual font bytes"));
                }
                total = total
                    .checked_add(resource.data.len())
                    .ok_or_else(|| limit("total font bytes"))?;
                if total > MAX_TOTAL_FONT_BYTES {
                    return Err(limit("total font bytes"));
                }
            }
        }
    }
    Ok(())
}

fn finish_font(font: &EmbeddedFont) -> Result<()> {
    if font.typeface.is_empty() {
        return Err(invalid("embeddedFont is missing its font descriptor"));
    }
    if font.faces.is_empty() {
        return Err(invalid("embeddedFont has no font-data face"));
    }
    if font.faces.len() > 4 {
        return Err(invalid("embeddedFont has more than four styles"));
    }
    for pair in font.faces.windows(2) {
        if pair[0].style.rank() >= pair[1].style.rank() {
            return Err(invalid(
                "embedded-font styles are duplicated or out of schema order",
            ));
        }
    }
    Ok(())
}

fn validate_font_relationship_sources(package: &OpcPackage, presentation: &str) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_font_relationship(relationship.reltype()))
    {
        return Err(invalid("package root cannot source font relationships"));
    }
    for part in package.iter_parts() {
        if part.partname().as_str() != presentation
            && part
                .rels()
                .iter()
                .any(|relationship| is_font_relationship(relationship.reltype()))
        {
            return Err(invalid(format!(
                "font relationship has invalid source '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn validate_inbound_font_graph(
    package: &OpcPackage,
    presentation_name: &str,
    presentation: &dyn Part,
    references: &HashSet<String>,
    targets: &HashSet<String>,
) -> Result<()> {
    for relationship in presentation
        .rels()
        .iter()
        .filter(|relationship| is_font_relationship(relationship.reltype()))
    {
        if !references.contains(relationship.r_id()) {
            return Err(invalid(format!(
                "unreferenced font relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            let target = relationship.target_partname()?;
            if targets.contains(target.as_str())
                && is_font_relationship(relationship.reltype())
                && (part.partname().as_str() != presentation_name
                    || !is_font_relationship(relationship.reltype())
                    || !references.contains(relationship.r_id()))
            {
                return Err(invalid(format!(
                    "font part '{target}' has an invalid inbound relationship"
                )));
            }
        }
    }
    Ok(())
}

fn reject_orphan_font_parts(package: &OpcPackage, targets: &HashSet<String>) -> Result<()> {
    for part in package.iter_parts() {
        if (part.partname().as_str().starts_with("/ppt/fonts/")
            || is_font_content_type(part.content_type()))
            && !targets.contains(part.partname().as_str())
            && !part_is_referenced(package, part.partname())?
        {
            return Err(invalid(format!("orphan font part '{}'", part.partname())));
        }
    }
    Ok(())
}

fn require_presentation(part: &dyn Part) -> Result<()> {
    if matches!(
        part.content_type(),
        PRESENTATION_CT | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main part has unsupported presentation content type '{}'",
            part.content_type()
        )))
    }
}

fn metadata_only(value: &PresentationEmbeddedFonts) -> PresentationEmbeddedFonts {
    let mut value = value.clone();
    for font in &mut value.fonts {
        for face in &mut font.faces {
            face.resource = None;
        }
    }
    value
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn has_inbound_outside_relationships(
    package: &OpcPackage,
    target: &PackURI,
    presentation: &PackURI,
    replaced_relationships: &HashSet<String>,
) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if relationship.target_partname()? == *target
                && (part.partname() != presentation
                    || !replaced_relationships.contains(relationship.r_id()))
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn parse_font_key(value: &str) -> Result<[u8; 16]> {
    let value = value
        .strip_prefix('{')
        .and_then(|value| value.strip_suffix('}'))
        .unwrap_or(value);
    if value.len() != 36
        || ![8, 13, 18, 23]
            .iter()
            .all(|offset| value.as_bytes()[*offset] == b'-')
    {
        return Err(invalid("font key must be a GUID"));
    }
    let digits = value
        .bytes()
        .filter(|byte| *byte != b'-')
        .collect::<Vec<_>>();
    if digits.len() != 32 || !digits.iter().all(u8::is_ascii_hexdigit) {
        return Err(invalid("font key must be a hexadecimal GUID"));
    }
    let mut key = [0u8; 16];
    for (index, output) in key.iter_mut().enumerate() {
        let pair = std::str::from_utf8(&digits[index * 2..index * 2 + 2]).map_err(xml_error)?;
        *output = u8::from_str_radix(pair, 16).map_err(xml_error)?;
    }
    Ok(key)
}

fn parse_panose(value: &str) -> Result<[u8; 10]> {
    if value.len() != 20 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid("panose must contain exactly 20 hexadecimal digits"));
    }
    let mut output = [0u8; 10];
    for (index, byte) in output.iter_mut().enumerate() {
        *byte = u8::from_str_radix(&value[index * 2..index * 2 + 2], 16).map_err(xml_error)?;
    }
    Ok(output)
}

fn hex_panose(value: &[u8; 10]) -> String {
    let mut output = String::with_capacity(20);
    for byte in value {
        use std::fmt::Write;
        write!(&mut output, "{byte:02X}").expect("writing to String cannot fail");
    }
    output
}

fn parse_u8(value: &str, name: &str) -> Result<u8> {
    value
        .parse()
        .map_err(|_| invalid(format!("invalid {name} byte value '{value}'")))
}
fn is_font_relationship(value: &str) -> bool {
    matches!(value, FONT_REL | STRICT_FONT_REL)
}
fn is_font_content_type(value: &str) -> bool {
    matches!(value, FONT_DATA_CT | FONT_TTF_CT)
}
fn bounded_string(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("embedded-font string bytes"))
    }
}
fn add_string_bytes(total: &mut usize, count: usize) -> Result<()> {
    *total = total
        .checked_add(count)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
fn validate_relationship_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID cannot be empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return Err(invalid(format!("invalid relationship ID '{value}'")));
    }
    Ok(())
}
fn resolved_namespace(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
fn attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'\"');
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
    invalid(format!("embedded-font {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn root() -> std::path::PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
    }
    fn package(conformance: EmbeddedFontConformance) -> OpcPackage {
        let mut package = OpcPackage::new();
        let uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let xml = format!(
            "<p:presentation xmlns:p=\"{}\"><p:sldMasterIdLst/><p:defaultTextStyle/></p:presentation>",
            conformance.pml()
        );
        package.add_part(Box::new(BlobPart::new(
            uri,
            PRESENTATION_CT.into(),
            xml.into_bytes(),
        )));
        let office_rel = match conformance {
            EmbeddedFontConformance::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            },
            EmbeddedFontConformance::Strict => {
                "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
            },
        };
        package.rels_mut().add_relationship(
            office_rel.into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        package
    }
    fn value() -> PresentationEmbeddedFonts {
        PresentationEmbeddedFonts {
            fonts: vec![EmbeddedFont {
                typeface: "A&B".into(),
                panose: Some([2, 11, 6, 4, 2, 2, 2, 2, 2, 4]),
                pitch_family: Some(34),
                charset: Some(0),
                faces: vec![
                    EmbeddedFontFace {
                        style: EmbeddedFontStyle::Regular,
                        relationship_id: "rIdFont1".into(),
                        resource: Some(EmbeddedFontResource {
                            part_name: "/ppt/fonts/font1.fntdata".into(),
                            content_type: FONT_DATA_CT.into(),
                            data: vec![0, 1, 2, 3],
                        }),
                    },
                    EmbeddedFontFace {
                        style: EmbeddedFontStyle::BoldItalic,
                        relationship_id: "rIdFont2".into(),
                        resource: Some(EmbeddedFontResource {
                            part_name: "/ppt/fonts/font2.fntdata".into(),
                            content_type: FONT_DATA_CT.into(),
                            data: vec![4, 5, 6],
                        }),
                    },
                ],
            }],
        }
    }

    #[test]
    fn strict_xml_round_trip_and_mce_fallback() {
        let expected = value();
        let fragment =
            write_embedded_font_list(&expected, EmbeddedFontConformance::Strict).unwrap();
        let xml = [
            format!("<p:presentation xmlns:p=\"{STRICT_PML}\">").as_bytes(),
            fragment.as_slice(),
            b"</p:presentation>",
        ]
        .concat();
        let parsed = parse_embedded_fonts(&xml).unwrap().unwrap();
        assert_eq!(parsed.fonts[0].typeface, "A&B");
        assert!(
            parsed.fonts[0]
                .faces
                .iter()
                .all(|face| face.resource.is_none())
        );
        let mce = format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:AlternateContent><mc:Choice Requires="x"><x:future/></mc:Choice><mc:Fallback><p:embeddedFontLst><p:embeddedFont><p:font typeface="Fallback"/><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></mc:Fallback></mc:AlternateContent></p:presentation>"#
        );
        assert_eq!(
            parse_embedded_fonts(mce.as_bytes()).unwrap().unwrap().fonts[0].typeface,
            "Fallback"
        );
    }

    #[test]
    fn loads_libreoffice_and_poi_reference_packages() {
        let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(
            root().join("test-data/libreoffice-core/sd/qa/unit/data/BoldonseFontEmbedded.pptx"),
        )
        .unwrap();
        let mut libreoffice = package(EmbeddedFontConformance::Transitional);
        let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
        libreoffice
            .get_part_mut(&presentation_uri)
            .unwrap()
            .set_blob(physical.blob_for(&presentation_uri).unwrap());
        let font_uri = PackURI::new("/ppt/fonts/font1.fntdata").unwrap();
        libreoffice.add_part(Box::new(BlobPart::new(
            font_uri.clone(),
            FONT_DATA_CT.into(),
            physical.blob_for(&font_uri).unwrap(),
        )));
        libreoffice
            .get_part_mut(&presentation_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                FONT_REL.into(),
                "fonts/font1.fntdata".into(),
                "rId3".into(),
                false,
            );
        let fonts = load_embedded_fonts(&libreoffice).unwrap().unwrap();
        assert_eq!(fonts.fonts[0].typeface, "Boldonse");
        assert_eq!(
            fonts.fonts[0].faces[0]
                .resource
                .as_ref()
                .unwrap()
                .data
                .len(),
            36_187
        );
        let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(
            root().join("test-data/poi/test-data/slideshow/placeholder-layout-color.pptx"),
        )
        .unwrap();
        let mut poi = package(EmbeddedFontConformance::Transitional);
        let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
        poi.get_part_mut(&presentation_uri)
            .unwrap()
            .set_blob(physical.blob_for(&presentation_uri).unwrap());
        for (index, relationship_id) in
            (1..=6).zip(["rId4", "rId5", "rId6", "rId7", "rId8", "rId9"])
        {
            let uri = PackURI::new(format!("/ppt/fonts/font{index}.fntdata")).unwrap();
            let data = physical.blob_for(&uri).unwrap();
            poi.add_part(Box::new(BlobPart::new(uri, FONT_DATA_CT.into(), data)));
            poi.get_part_mut(&presentation_uri)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    FONT_REL.into(),
                    format!("fonts/font{index}.fntdata"),
                    relationship_id.into(),
                    false,
                );
        }
        let fonts = load_embedded_fonts(&poi).unwrap().unwrap();
        assert_eq!(fonts.fonts.len(), 3);
        let roboto = fonts
            .fonts
            .iter()
            .find(|font| font.typeface == "Roboto")
            .unwrap();
        assert_eq!(roboto.faces.len(), 4);
        assert!(roboto.faces.iter().all(|face| {
            face.resource
                .as_ref()
                .is_some_and(|resource| !resource.data.is_empty())
        }));
    }

    #[test]
    fn package_writer_round_trips_strict_graph_and_schema_position() {
        let mut package = package(EmbeddedFontConformance::Strict);
        let expected = value();
        store_embedded_fonts(&mut package, &expected, EmbeddedFontConformance::Strict).unwrap();
        assert_eq!(load_embedded_fonts(&package).unwrap().unwrap(), expected);
        let xml = package.main_document_part().unwrap().blob();
        let list = memchr::memmem::find(xml, b"<p:embeddedFontLst").unwrap();
        let defaults = memchr::memmem::find(xml, b"<p:defaultTextStyle").unwrap();
        assert!(list < defaults);
    }

    #[test]
    fn rejects_malformed_xml_duplicates_and_caps() {
        for xml in [
            format!(r#"<p:presentation xmlns:p="{PML}"/>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A" panose="12"/><p:regular xmlns:r="{REL_NS}" r:id="rId1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:bold xmlns:r="{REL_NS}" r:id="rId1"/><p:regular r:id="rId2"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{PML}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#),
            format!(r#"<!DOCTYPE x><p:presentation xmlns:p="{PML}"/>"#),
        ].into_iter().skip(1) { assert!(parse_embedded_fonts(xml.as_bytes()).is_err(), "{xml}"); }
        assert!(parse_embedded_fonts(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
        let duplicate = PresentationEmbeddedFonts {
            fonts: vec![value().fonts[0].clone(), value().fonts[0].clone()],
        };
        assert!(
            write_embedded_font_list(&duplicate, EmbeddedFontConformance::Transitional).is_err()
        );
    }

    #[test]
    fn rejects_external_orphan_and_outbound_graphs() {
        let mut external = package(EmbeddedFontConformance::Transitional);
        let xml = format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL_NS}"><p:embeddedFontLst><p:embeddedFont><p:font typeface="A"/><p:regular r:id="rIdFont1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>"#
        );
        external
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .set_blob(xml.into_bytes());
        external
            .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                FONT_REL.into(),
                "https://invalid.example/font".into(),
                "rIdFont1".into(),
                true,
            );
        assert!(load_embedded_fonts(&external).is_err());

        let mut orphan = package(EmbeddedFontConformance::Transitional);
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/fonts/orphan.fntdata").unwrap(),
            FONT_DATA_CT.into(),
            vec![1],
        )));
        assert!(load_embedded_fonts(&orphan).is_err());

        let mut outbound = package(EmbeddedFontConformance::Transitional);
        store_embedded_fonts(
            &mut outbound,
            &value(),
            EmbeddedFontConformance::Transitional,
        )
        .unwrap();
        outbound
            .get_part_mut(&PackURI::new("/ppt/fonts/font1.fntdata").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "other.bin".into(),
                "rId1".into(),
                false,
            );
        assert!(load_embedded_fonts(&outbound).is_err());
    }

    #[test]
    fn guid_obfuscation_is_reversible_and_fs_type_is_validated() {
        let mut data = (0u8..64).collect::<Vec<_>>();
        let original = data.clone();
        let key = "{00010203-0405-0607-0809-0A0B0C0D0E0F}";
        obfuscate_embedded_font_data(&mut data, key).unwrap();
        assert_ne!(data, original);
        assert_eq!(data[0], original[0] ^ 0x0F);
        assert_eq!(data[16], original[16] ^ 0x0F);
        deobfuscate_embedded_font_data(&mut data, key).unwrap();
        assert_eq!(data, original);
        assert!(obfuscate_embedded_font_data(&mut [0; 31], key).is_err());
        assert!(obfuscate_embedded_font_data(&mut [0; 32], "bad").is_err());

        assert!(
            EmbeddedFontLicensing::from_fs_type(0)
                .unwrap()
                .installable()
        );
        let editable = EmbeddedFontLicensing::from_fs_type(0x0108).unwrap();
        assert!(editable.editable && editable.no_subsetting && !editable.installable());
        assert!(EmbeddedFontLicensing::from_fs_type(0x0006).is_err());
        assert!(EmbeddedFontLicensing::from_fs_type(0x8000).is_err());
    }

    #[test]
    fn generated_crud_allocates_collisions_and_preserves_unknown_xml_atomically() {
        let mut package = package(EmbeddedFontConformance::Transitional);
        let presentation_uri = PackURI::new("/ppt/presentation.xml").unwrap();
        let original = package.get_part(&presentation_uri).unwrap().blob();
        let marker = memchr::memmem::find(original, b"<p:defaultTextStyle").unwrap();
        let mut xml = original.to_vec();
        xml.splice(marker..marker, b"<!--font-preserve-->".iter().copied());
        package
            .get_part_mut(&presentation_uri)
            .unwrap()
            .set_blob(xml);

        let generated = EmbeddedFont {
            typeface: "Generated".into(),
            panose: Some([2, 11, 6, 4, 2, 2, 2, 2, 2, 4]),
            pitch_family: Some(34),
            charset: Some(0),
            faces: vec![EmbeddedFontFace {
                style: EmbeddedFontStyle::Regular,
                relationship_id: String::new(),
                resource: Some(EmbeddedFontResource {
                    part_name: String::new(),
                    content_type: String::new(),
                    data: vec![7; 64],
                }),
            }],
        };
        add_embedded_font(
            &mut package,
            generated,
            EmbeddedFontConformance::Transitional,
        )
        .unwrap();
        let found = find_embedded_font(&package, "generated").unwrap().unwrap();
        assert_eq!(found.faces[0].relationship_id, "rIdFont1");
        assert_eq!(
            found.faces[0].resource.as_ref().unwrap().part_name,
            "/ppt/fonts/font1.fntdata"
        );
        assert!(
            package
                .get_part(&presentation_uri)
                .unwrap()
                .blob()
                .windows(b"<!--font-preserve-->".len())
                .any(|window| window == b"<!--font-preserve-->")
        );

        let before = package.get_part(&presentation_uri).unwrap().blob().to_vec();
        let parts = package.part_count();
        let duplicate = found.clone();
        assert!(
            add_embedded_font(
                &mut package,
                duplicate,
                EmbeddedFontConformance::Transitional
            )
            .is_err()
        );
        assert_eq!(package.get_part(&presentation_uri).unwrap().blob(), before);
        assert_eq!(package.part_count(), parts);

        let mut replacement = found;
        replacement.typeface = "Renamed".into();
        replacement.charset = Some(1);
        update_embedded_font(
            &mut package,
            "Generated",
            replacement,
            EmbeddedFontConformance::Transitional,
        )
        .unwrap();
        assert!(find_embedded_font(&package, "Renamed").unwrap().is_some());
        assert!(
            remove_embedded_font(
                &mut package,
                "Renamed",
                EmbeddedFontConformance::Transitional
            )
            .unwrap()
        );
        assert!(load_embedded_fonts(&package).unwrap().is_none());
    }

    #[test]
    fn shared_font_parts_survive_face_and_external_owner_removal() {
        let mut package = package(EmbeddedFontConformance::Transitional);
        let shared = EmbeddedFontResource {
            part_name: "/ppt/fonts/shared.fntdata".into(),
            content_type: FONT_DATA_CT.into(),
            data: vec![3; 64],
        };
        let graph = PresentationEmbeddedFonts {
            fonts: vec![
                EmbeddedFont {
                    typeface: "First".into(),
                    panose: None,
                    pitch_family: None,
                    charset: None,
                    faces: vec![EmbeddedFontFace {
                        style: EmbeddedFontStyle::Regular,
                        relationship_id: "rIdFontA".into(),
                        resource: Some(shared.clone()),
                    }],
                },
                EmbeddedFont {
                    typeface: "Second".into(),
                    panose: None,
                    pitch_family: None,
                    charset: None,
                    faces: vec![EmbeddedFontFace {
                        style: EmbeddedFontStyle::Regular,
                        relationship_id: "rIdFontB".into(),
                        resource: Some(shared),
                    }],
                },
            ],
        };
        store_embedded_fonts(&mut package, &graph, EmbeddedFontConformance::Transitional).unwrap();
        assert_eq!(
            load_embedded_fonts(&package).unwrap().unwrap().fonts.len(),
            2
        );
        remove_embedded_font(&mut package, "First", EmbeddedFontConformance::Transitional).unwrap();
        let font_uri = PackURI::new("/ppt/fonts/shared.fntdata").unwrap();
        assert!(package.contains_part(&font_uri));

        let owner_uri = PackURI::new("/ppt/unknown-owner.bin").unwrap();
        let mut owner = BlobPart::new(
            owner_uri.clone(),
            "application/octet-stream".into(),
            vec![1],
        );
        owner.rels_mut().add_relationship(
            "urn:shared-resource".into(),
            "fonts/shared.fntdata".into(),
            "rIdShared".into(),
            false,
        );
        package.add_part(Box::new(owner));
        remove_embedded_font(
            &mut package,
            "Second",
            EmbeddedFontConformance::Transitional,
        )
        .unwrap();
        assert!(package.contains_part(&font_uri));
        assert!(load_embedded_fonts(&package).unwrap().is_none());
    }
}
