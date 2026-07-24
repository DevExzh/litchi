//! Typed PresentationML audio/video pictures and inert package media resources.

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashSet};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const STRICT_AUDIO_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/audio";
const STRICT_VIDEO_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/video";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const MEDIA_EXTENSION_URI: &str = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_NODES: usize = 500_000;
const MAX_DEPTH: usize = 256;
const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
const MAX_MEDIA: usize = 1024;
const MAX_BOOKMARKS: usize = 4096;
const MAX_PAYLOAD_BYTES: usize = 128 * 1024 * 1024;
const MAX_TOTAL_PAYLOAD_BYTES: usize = 512 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SlideMediaConformance {
    Transitional,
    Strict,
}

impl SlideMediaConformance {
    fn pml(self) -> &'static str {
        match self {
            Self::Transitional => PML,
            Self::Strict => STRICT_PML,
        }
    }
    fn dml(self) -> &'static str {
        match self {
            Self::Transitional => DML,
            Self::Strict => STRICT_DML,
        }
    }
    fn rel(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }
    fn image_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::IMAGE,
            Self::Strict => rt::STRICT_IMAGE,
        }
    }
    fn media_rel(self, kind: SlideMediaKind) -> &'static str {
        match (self, kind) {
            (Self::Transitional, SlideMediaKind::Audio) => rt::AUDIO,
            (Self::Transitional, SlideMediaKind::Video) => rt::VIDEO,
            (Self::Strict, SlideMediaKind::Audio) => STRICT_AUDIO_REL,
            (Self::Strict, SlideMediaKind::Video) => STRICT_VIDEO_REL,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SlideMediaKind {
    Audio,
    Video,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaResource {
    pub part_name: String,
    pub content_type: String,
    /// Stored and returned verbatim. The payload is never decoded or executed.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideMediaPoster {
    pub relationship_id: String,
    pub resource: Option<MediaResource>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideMediaTransform {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaTrim {
    pub start: Option<String>,
    pub end: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaFade {
    pub fade_in: Option<String>,
    pub fade_out: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaBookmark {
    pub name: Option<String>,
    pub time: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OfficeMediaExtension {
    pub embed_relationship_id: Option<String>,
    pub link_relationship_id: Option<String>,
    pub trim: Option<MediaTrim>,
    pub fade: Option<MediaFade>,
    pub bookmarks: Vec<MediaBookmark>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideMediaPicture {
    pub shape_id: u32,
    pub name: String,
    pub kind: SlideMediaKind,
    /// The ISO/IEC 29500 `a:audioFile` or `a:videoFile` relationship identifier.
    pub relationship_id: String,
    /// Filled by package loading and required by package storage.
    pub resource: Option<MediaResource>,
    pub poster: Option<SlideMediaPoster>,
    pub transform: Option<SlideMediaTransform>,
    pub office_extension: Option<OfficeMediaExtension>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SlideMediaList {
    pub pictures: Vec<SlideMediaPicture>,
}

#[derive(Clone)]
struct Attribute {
    namespace: String,
    name: String,
    value: String,
}
#[derive(Clone)]
struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}

/// Parses all audio/video `p:pic` elements from a complete Slide part.
pub fn parse_slide_media(xml: &[u8]) -> Result<SlideMediaList> {
    let root = parse_document(xml)?;
    let conformance = conformance(&root)?;
    let mut pictures = Vec::new();
    collect_pictures(&root, conformance, &mut pictures)?;
    let value = SlideMediaList { pictures };
    validate_value(&value, false)?;
    Ok(value)
}

fn collect_pictures(
    node: &Node,
    conformance: SlideMediaConformance,
    output: &mut Vec<SlideMediaPicture>,
) -> Result<()> {
    if node.namespace == conformance.pml() && node.name == "pic" {
        if let Some(picture) = parse_picture(node, conformance)? {
            if output.len() == MAX_MEDIA {
                return Err(limit("media count"));
            }
            output.push(picture);
        }
    }
    for child in &node.children {
        collect_pictures(child, conformance, output)?;
    }
    Ok(())
}

fn parse_picture(
    node: &Node,
    conformance: SlideMediaConformance,
) -> Result<Option<SlideMediaPicture>> {
    let Some(nv_pic) = one_child(node, conformance.pml(), "nvPicPr")? else {
        return Ok(None);
    };
    let Some(nv_pr) = one_child(nv_pic, conformance.pml(), "nvPr")? else {
        return Ok(None);
    };
    let audio = one_child(nv_pr, conformance.dml(), "audioFile")?;
    let video = one_child(nv_pr, conformance.dml(), "videoFile")?;
    let (kind, media) = match (audio, video) {
        (Some(_), Some(_)) => {
            return Err(invalid(
                "media picture contains both audioFile and videoFile",
            ));
        },
        (Some(value), None) => (SlideMediaKind::Audio, value),
        (None, Some(value)) => (SlideMediaKind::Video, value),
        (None, None) => return Ok(None),
    };
    leaf(media, "media file")?;
    let relationship_id = required(media, conformance.rel(), "link")?.to_owned();
    no_attributes(media, &[(conformance.rel(), "link")])?;
    let c_nv_pr = required_child(nv_pic, conformance.pml(), "cNvPr")?;
    let shape_id = required(c_nv_pr, "", "id")?
        .parse()
        .map_err(|_| invalid("invalid media shape id"))?;
    let name = optional(c_nv_pr, "", "name").unwrap_or_default().to_owned();
    let poster = parse_poster(node, conformance)?;
    let transform = parse_transform(node, conformance)?;
    let office_extension = find_office_media(nv_pr, conformance)?
        .map(|media| parse_office_media(media, conformance))
        .transpose()?;
    Ok(Some(SlideMediaPicture {
        shape_id,
        name,
        kind,
        relationship_id,
        resource: None,
        poster,
        transform,
        office_extension,
    }))
}

fn parse_poster(
    node: &Node,
    conformance: SlideMediaConformance,
) -> Result<Option<SlideMediaPoster>> {
    let Some(fill) = one_child(node, conformance.pml(), "blipFill")? else {
        return Ok(None);
    };
    let Some(blip) = one_child(fill, conformance.dml(), "blip")? else {
        return Ok(None);
    };
    let relationship_id = required(blip, conformance.rel(), "embed")?.to_owned();
    Ok(Some(SlideMediaPoster {
        relationship_id,
        resource: None,
    }))
}

fn parse_transform(
    node: &Node,
    conformance: SlideMediaConformance,
) -> Result<Option<SlideMediaTransform>> {
    let Some(properties) = one_child(node, conformance.pml(), "spPr")? else {
        return Ok(None);
    };
    let Some(transform) = one_child(properties, conformance.dml(), "xfrm")? else {
        return Ok(None);
    };
    let offset = required_child(transform, conformance.dml(), "off")?;
    let extent = required_child(transform, conformance.dml(), "ext")?;
    leaf(offset, "transform offset")?;
    leaf(extent, "transform extent")?;
    Ok(Some(SlideMediaTransform {
        x: parse_i64(required(offset, "", "x")?, "x")?,
        y: parse_i64(required(offset, "", "y")?, "y")?,
        width: parse_i64(required(extent, "", "cx")?, "width")?,
        height: parse_i64(required(extent, "", "cy")?, "height")?,
    }))
}

fn find_office_media<'a>(
    node: &'a Node,
    conformance: SlideMediaConformance,
) -> Result<Option<&'a Node>> {
    let mut found = None;
    for list in node
        .children
        .iter()
        .filter(|child| child.namespace == conformance.pml() && child.name == "extLst")
    {
        for extension in list
            .children
            .iter()
            .filter(|child| child.namespace == conformance.pml() && child.name == "ext")
        {
            for media in extension
                .children
                .iter()
                .filter(|child| child.namespace == P14 && child.name == "media")
            {
                if found.replace(media).is_some() {
                    return Err(invalid("media picture has multiple p14:media extensions"));
                }
            }
        }
    }
    Ok(found)
}

fn parse_office_media(
    node: &Node,
    conformance: SlideMediaConformance,
) -> Result<OfficeMediaExtension> {
    whitespace(node)?;
    let embed_relationship_id = optional(node, conformance.rel(), "embed").map(str::to_owned);
    let link_relationship_id = optional(node, conformance.rel(), "link").map(str::to_owned);
    no_attributes(
        node,
        &[(conformance.rel(), "embed"), (conformance.rel(), "link")],
    )?;
    if embed_relationship_id.is_none() && link_relationship_id.is_none() {
        return Err(invalid("p14:media requires r:embed or r:link"));
    }
    let trim_node = one_child(node, P14, "trim")?;
    let fade_node = one_child(node, P14, "fade")?;
    let bookmarks_node = one_child(node, P14, "bmkLst")?;
    for child in &node.children {
        if child.namespace == P14
            && !matches!(child.name.as_str(), "trim" | "fade" | "bmkLst")
            && !(child.namespace == conformance.pml() && child.name == "extLst")
        {
            return Err(invalid(format!(
                "unsupported p14:media child '{}'",
                child.name
            )));
        }
    }
    let trim = trim_node.map(|value| parse_trim(value)).transpose()?;
    let fade = fade_node.map(|value| parse_fade(value)).transpose()?;
    let mut bookmarks = Vec::new();
    if let Some(list) = bookmarks_node {
        whitespace(list)?;
        no_attributes(list, &[])?;
        if list.children.len() > MAX_BOOKMARKS {
            return Err(limit("bookmark count"));
        }
        for child in &list.children {
            require_node(child, P14, "bmk")?;
            leaf(child, "media bookmark")?;
            bookmarks.push(MediaBookmark {
                name: optional(child, "", "name").map(str::to_owned),
                time: optional(child, "", "time").map(str::to_owned),
            });
            no_attributes(child, &[("", "name"), ("", "time")])?;
        }
    }
    Ok(OfficeMediaExtension {
        embed_relationship_id,
        link_relationship_id,
        trim,
        fade,
        bookmarks,
    })
}

fn parse_trim(node: &Node) -> Result<MediaTrim> {
    leaf(node, "media trim")?;
    no_attributes(node, &[("", "st"), ("", "end")])?;
    Ok(MediaTrim {
        start: optional(node, "", "st").map(str::to_owned),
        end: optional(node, "", "end").map(str::to_owned),
    })
}

fn parse_fade(node: &Node) -> Result<MediaFade> {
    leaf(node, "media fade")?;
    no_attributes(node, &[("", "in"), ("", "out")])?;
    Ok(MediaFade {
        fade_in: optional(node, "", "in").map(str::to_owned),
        fade_out: optional(node, "", "out").map(str::to_owned),
    })
}

/// Deterministically serializes self-contained `p:pic` fragments.
pub fn write_slide_media_pictures(
    value: &SlideMediaList,
    conformance: SlideMediaConformance,
) -> Result<Vec<u8>> {
    validate_value(value, false)?;
    let mut output = Vec::new();
    for picture in &value.pictures {
        write_picture(&mut output, picture, conformance);
    }
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized XML bytes"));
    }
    Ok(output)
}

fn write_picture(
    output: &mut Vec<u8>,
    picture: &SlideMediaPicture,
    conformance: SlideMediaConformance,
) {
    output.extend_from_slice(b"<p:pic xmlns:p=\"");
    escape(output, conformance.pml());
    output.extend_from_slice(b"\" xmlns:a=\"");
    escape(output, conformance.dml());
    output.extend_from_slice(b"\" xmlns:r=\"");
    escape(output, conformance.rel());
    if picture.office_extension.is_some() {
        output.extend_from_slice(b"\" xmlns:p14=\"");
        escape(output, P14);
    }
    output.extend_from_slice(b"\"><p:nvPicPr><p:cNvPr");
    attr(output, "id", &picture.shape_id.to_string());
    attr(output, "name", &picture.name);
    output.extend_from_slice(
        b"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr>",
    );
    output.extend_from_slice(match picture.kind {
        SlideMediaKind::Audio => b"<a:audioFile".as_slice(),
        SlideMediaKind::Video => b"<a:videoFile".as_slice(),
    });
    attr(output, "r:link", &picture.relationship_id);
    output.extend_from_slice(b"/>");
    if let Some(extension) = &picture.office_extension {
        write_office_extension(output, extension);
    }
    output.extend_from_slice(b"</p:nvPr></p:nvPicPr><p:blipFill>");
    if let Some(poster) = &picture.poster {
        output.extend_from_slice(b"<a:blip");
        attr(output, "r:embed", &poster.relationship_id);
        output.extend_from_slice(b"/><a:stretch><a:fillRect/></a:stretch>");
    }
    output.extend_from_slice(b"</p:blipFill><p:spPr>");
    if let Some(transform) = picture.transform {
        output.extend_from_slice(b"<a:xfrm><a:off");
        attr(output, "x", &transform.x.to_string());
        attr(output, "y", &transform.y.to_string());
        output.extend_from_slice(b"/><a:ext");
        attr(output, "cx", &transform.width.to_string());
        attr(output, "cy", &transform.height.to_string());
        output.extend_from_slice(b"/></a:xfrm>");
    }
    output.extend_from_slice(b"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>");
}

fn write_office_extension(output: &mut Vec<u8>, value: &OfficeMediaExtension) {
    output.extend_from_slice(b"<p:extLst><p:ext uri=\"");
    escape(output, MEDIA_EXTENSION_URI);
    output.extend_from_slice(b"\"><p14:media");
    if let Some(id) = &value.embed_relationship_id {
        attr(output, "r:embed", id);
    }
    if let Some(id) = &value.link_relationship_id {
        attr(output, "r:link", id);
    }
    if value.trim.is_none() && value.fade.is_none() && value.bookmarks.is_empty() {
        output.extend_from_slice(b"/></p:ext></p:extLst>");
        return;
    }
    output.push(b'>');
    if let Some(trim) = &value.trim {
        output.extend_from_slice(b"<p14:trim");
        if let Some(v) = &trim.start {
            attr(output, "st", v);
        }
        if let Some(v) = &trim.end {
            attr(output, "end", v);
        }
        output.extend_from_slice(b"/>");
    }
    if let Some(fade) = &value.fade {
        output.extend_from_slice(b"<p14:fade");
        if let Some(v) = &fade.fade_in {
            attr(output, "in", v);
        }
        if let Some(v) = &fade.fade_out {
            attr(output, "out", v);
        }
        output.extend_from_slice(b"/>");
    }
    if !value.bookmarks.is_empty() {
        output.extend_from_slice(b"<p14:bmkLst>");
        for bookmark in &value.bookmarks {
            output.extend_from_slice(b"<p14:bmk");
            if let Some(v) = &bookmark.name {
                attr(output, "name", v);
            }
            if let Some(v) = &bookmark.time {
                attr(output, "time", v);
            }
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</p14:bmkLst>");
    }
    output.extend_from_slice(b"</p14:media></p:ext></p:extLst>");
}

/// Loads media pictures and validates their complete internal OPC resource graph.
pub fn load_slide_media(package: &OpcPackage, slide_name: &PackURI) -> Result<SlideMediaList> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_media_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "package root cannot source slide-media relationships",
        ));
    }
    let slide = package.get_part(slide_name)?;
    require_slide(slide)?;
    let mut value = parse_slide_media(slide.blob())?;
    let conformance = conformance(&parse_document(slide.blob())?)?;
    let mut total = 0usize;
    let mut loaded: BTreeMap<String, MediaResource> = BTreeMap::new();
    for picture in &mut value.pictures {
        let target = relationship_target(
            slide,
            &picture.relationship_id,
            conformance.media_rel(picture.kind),
        )?;
        let resource = load_resource(
            package,
            &target,
            picture.kind,
            false,
            &mut total,
            &mut loaded,
        )?;
        picture.resource = Some(resource.clone());
        if let Some(extension) = &picture.office_extension {
            for id in extension
                .embed_relationship_id
                .iter()
                .chain(extension.link_relationship_id.iter())
            {
                let extension_target = relationship_target(slide, id, rt::MEDIA)?;
                if extension_target != target {
                    return Err(invalid(format!(
                        "p14 media relationship '{id}' does not target the ISO media resource"
                    )));
                }
            }
        }
        if let Some(poster) = picture.poster.as_mut() {
            let target =
                relationship_target(slide, &poster.relationship_id, conformance.image_rel())?;
            poster.resource = Some(load_resource(
                package,
                &target,
                picture.kind,
                true,
                &mut total,
                &mut loaded,
            )?);
        }
    }
    Ok(value)
}

fn relationship_target(part: &dyn Part, id: &str, expected_type: &str) -> Result<PackURI> {
    let relationship = part
        .rels()
        .get(id)
        .ok_or_else(|| invalid(format!("missing slide-media relationship '{id}'")))?;
    if relationship.reltype() != expected_type {
        return Err(invalid(format!(
            "relationship '{id}' has type '{}', expected '{expected_type}'",
            relationship.reltype()
        )));
    }
    if relationship.is_external() {
        return Err(invalid(format!(
            "external slide-media relationship '{id}' is not fetched"
        )));
    }
    relationship.target_partname().map_err(OoxmlError::Opc)
}

fn load_resource(
    package: &OpcPackage,
    target: &PackURI,
    kind: SlideMediaKind,
    image: bool,
    total: &mut usize,
    loaded: &mut BTreeMap<String, MediaResource>,
) -> Result<MediaResource> {
    if let Some(value) = loaded.get(target.as_str()) {
        return Ok(value.clone());
    }
    if !target.as_str().starts_with("/ppt/media/") {
        return Err(invalid(format!(
            "slide media resource '{target}' is outside /ppt/media"
        )));
    }
    let part = package.get_part(target)?;
    if !part.rels().is_empty() {
        return Err(invalid(format!(
            "slide media resource '{target}' has forbidden outbound relationships"
        )));
    }
    if image {
        if !is_image_content_type(part.content_type()) {
            return Err(invalid(format!(
                "poster '{target}' has non-image content type '{}'",
                part.content_type()
            )));
        }
    } else if !is_media_content_type(part.content_type(), kind) {
        return Err(invalid(format!(
            "media '{target}' has content type '{}' inconsistent with its media kind",
            part.content_type()
        )));
    }
    add_payload(total, part.blob().len())?;
    let value = MediaResource {
        part_name: target.to_string(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
    };
    loaded.insert(value.part_name.clone(), value.clone());
    Ok(value)
}

/// Adds a new set of media pictures and their inert internal resources to a Slide part.
pub fn store_slide_media(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    value: &SlideMediaList,
    conformance: SlideMediaConformance,
) -> Result<()> {
    validate_value(value, true)?;
    let slide = package.get_part(slide_name)?;
    require_slide(slide)?;
    if !parse_slide_media(slide.blob())?.pictures.is_empty() {
        return Err(invalid("slide already contains media pictures"));
    }
    if crate_conformance(&parse_document(slide.blob())?)? != conformance {
        return Err(invalid(
            "requested conformance does not match slide namespace",
        ));
    }
    let fragment = write_slide_media_pictures(value, conformance)?;
    let updated = insert_pictures(slide.blob(), &fragment, conformance)?;
    let mut relationships: BTreeMap<String, (String, String)> = BTreeMap::new();
    let mut parts: BTreeMap<String, MediaResource> = BTreeMap::new();
    for picture in &value.pictures {
        let resource = picture
            .resource
            .as_ref()
            .ok_or_else(|| invalid("media resource is required for package storage"))?;
        let uri = resource_uri(resource, false, Some(picture.kind))?;
        add_part_plan(package, &mut parts, resource)?;
        let target = uri.relative_ref(slide_name.base_uri());
        add_relationship_plan(
            &mut relationships,
            &picture.relationship_id,
            conformance.media_rel(picture.kind),
            &target,
        )?;
        if let Some(extension) = &picture.office_extension {
            for id in extension
                .embed_relationship_id
                .iter()
                .chain(extension.link_relationship_id.iter())
            {
                add_relationship_plan(&mut relationships, id, rt::MEDIA, &target)?;
            }
        }
        if let Some(poster) = &picture.poster {
            let resource = poster
                .resource
                .as_ref()
                .ok_or_else(|| invalid("poster resource is required for package storage"))?;
            let uri = resource_uri(resource, true, None)?;
            add_part_plan(package, &mut parts, resource)?;
            add_relationship_plan(
                &mut relationships,
                &poster.relationship_id,
                conformance.image_rel(),
                &uri.relative_ref(slide_name.base_uri()),
            )?;
        }
    }
    for id in relationships.keys() {
        if slide.rels().get(id).is_some() {
            return Err(invalid(format!(
                "slide relationship ID '{id}' already exists"
            )));
        }
    }
    package.get_part_mut(slide_name)?.set_blob(updated);
    for resource in parts.into_values() {
        let uri = PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?;
        package.add_part(Box::new(BlobPart::new(
            uri,
            resource.content_type,
            resource.data,
        )));
    }
    for (id, (relationship_type, target)) in relationships {
        package
            .get_part_mut(slide_name)?
            .rels_mut()
            .add_relationship(relationship_type, target, id, false);
    }
    Ok(())
}

fn insert_pictures(
    xml: &[u8],
    fragment: &[u8],
    conformance: SlideMediaConformance,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut sp_tree_depth = None;
    let mut position = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("slide XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value.as_ref() == conformance.pml().as_bytes());
                if depth == 0 && (!core || element.local_name().as_ref() != b"sld") {
                    return Err(invalid("slide root does not match conformance"));
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
                if core && element.local_name().as_ref() == b"spTree" {
                    if sp_tree_depth.replace(depth).is_some() {
                        return Err(invalid("slide has multiple shape trees"));
                    }
                }
            },
            Event::Empty(element) => {
                if element.local_name().as_ref() == b"spTree" {
                    return Err(invalid("cannot insert into an empty shape tree"));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected slide closing element"));
                }
                if sp_tree_depth == Some(depth) && element.local_name().as_ref() == b"spTree" {
                    position = Some(start);
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
    if depth != 0 {
        return Err(invalid("invalid slide XML"));
    }
    let position = position.ok_or_else(|| invalid("slide is missing a shape tree"))?;
    let size = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| limit("updated XML bytes"))?;
    if size > MAX_XML_BYTES {
        return Err(limit("updated XML bytes"));
    }
    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

fn validate_value(value: &SlideMediaList, require_resources: bool) -> Result<()> {
    if value.pictures.len() > MAX_MEDIA {
        return Err(limit("media count"));
    }
    let mut shape_ids = HashSet::new();
    let mut resources: BTreeMap<String, MediaResource> = BTreeMap::new();
    for picture in &value.pictures {
        if !(1..=67_098_623).contains(&picture.shape_id) {
            return Err(invalid(
                "media shape id is outside Office's supported range",
            ));
        }
        if !shape_ids.insert(picture.shape_id) {
            return Err(invalid(format!(
                "duplicate media shape id {}",
                picture.shape_id
            )));
        }
        bounded(&picture.name)?;
        validate_id(&picture.relationship_id)?;
        if let Some(transform) = picture.transform {
            if transform.width <= 0 || transform.height <= 0 {
                return Err(invalid("media transform width and height must be positive"));
            }
        }
        if require_resources && picture.resource.is_none() {
            return Err(invalid("media resource is required for package storage"));
        }
        if let Some(resource) = &picture.resource {
            resource_uri(resource, false, Some(picture.kind))?;
            merge_resource(&mut resources, resource)?;
        }
        if let Some(poster) = &picture.poster {
            validate_id(&poster.relationship_id)?;
            if require_resources && poster.resource.is_none() {
                return Err(invalid("poster resource is required for package storage"));
            }
            if let Some(resource) = &poster.resource {
                resource_uri(resource, true, None)?;
                merge_resource(&mut resources, resource)?;
            }
        }
        if let Some(extension) = &picture.office_extension {
            validate_extension(extension)?;
        }
    }
    let mut total = 0usize;
    for resource in resources.values() {
        add_payload(&mut total, resource.data.len())?;
    }
    Ok(())
}

fn validate_extension(value: &OfficeMediaExtension) -> Result<()> {
    if value.embed_relationship_id.is_none() && value.link_relationship_id.is_none() {
        return Err(invalid(
            "p14 media extension requires embed or link relationship",
        ));
    }
    for id in value
        .embed_relationship_id
        .iter()
        .chain(value.link_relationship_id.iter())
    {
        validate_id(id)?;
    }
    if let Some(trim) = &value.trim {
        for time in trim.start.iter().chain(trim.end.iter()) {
            validate_time(time)?;
        }
    }
    if let Some(fade) = &value.fade {
        for time in fade.fade_in.iter().chain(fade.fade_out.iter()) {
            validate_time(time)?;
        }
    }
    if value.bookmarks.len() > MAX_BOOKMARKS {
        return Err(limit("bookmark count"));
    }
    let mut names = HashSet::new();
    let mut times = HashSet::new();
    for bookmark in &value.bookmarks {
        if let Some(name) = &bookmark.name {
            bounded(name)?;
            if !names.insert(name) {
                return Err(invalid(format!("duplicate media bookmark name '{name}'")));
            }
        }
        if let Some(time) = &bookmark.time {
            validate_time(time)?;
            if !times.insert(time) {
                return Err(invalid(format!("duplicate media bookmark time '{time}'")));
            }
        }
    }
    Ok(())
}

fn validate_time(value: &str) -> Result<()> {
    bounded(value)?;
    let split = value
        .find(|character: char| !character.is_ascii_digit() && character != '.')
        .unwrap_or(value.len());
    let (number, unit) = value.split_at(split);
    let mut pieces = number.split('.');
    let whole = pieces.next().unwrap_or_default();
    let fraction = pieces.next();
    if whole.is_empty()
        || !whole.bytes().all(|b| b.is_ascii_digit())
        || fraction.is_some_and(|v| v.is_empty() || !v.bytes().all(|b| b.is_ascii_digit()))
        || pieces.next().is_some()
        || !matches!(unit, "" | "h" | "min" | "s" | "ms" | "µs" | "ns")
    {
        return Err(invalid(format!("invalid universal time offset '{value}'")));
    }
    Ok(())
}

fn resource_uri(
    resource: &MediaResource,
    image: bool,
    kind: Option<SlideMediaKind>,
) -> Result<PackURI> {
    let uri = PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?;
    if !uri.as_str().starts_with("/ppt/media/") {
        return Err(invalid(format!("resource '{uri}' is outside /ppt/media")));
    }
    if resource.content_type.is_empty()
        || resource
            .content_type
            .bytes()
            .any(|b| b.is_ascii_whitespace())
    {
        return Err(invalid("invalid media resource content type"));
    }
    if image {
        if !is_image_content_type(&resource.content_type) {
            return Err(invalid("poster resource has non-image content type"));
        }
    } else if !is_media_content_type(&resource.content_type, kind.expect("media kind")) {
        return Err(invalid(
            "media resource content type is inconsistent with its kind",
        ));
    }
    if resource.data.len() > MAX_PAYLOAD_BYTES {
        return Err(limit("individual payload bytes"));
    }
    Ok(uri)
}

fn merge_resource(
    resources: &mut BTreeMap<String, MediaResource>,
    resource: &MediaResource,
) -> Result<()> {
    if let Some(old) = resources.get(&resource.part_name) {
        if old != resource {
            return Err(invalid(format!(
                "conflicting resource part '{}'",
                resource.part_name
            )));
        }
    } else {
        resources.insert(resource.part_name.clone(), resource.clone());
    }
    Ok(())
}
fn add_part_plan(
    package: &OpcPackage,
    parts: &mut BTreeMap<String, MediaResource>,
    resource: &MediaResource,
) -> Result<()> {
    if package
        .iter_parts()
        .any(|part| part.partname().as_str() == resource.part_name)
    {
        return Err(invalid(format!(
            "resource part '{}' already exists",
            resource.part_name
        )));
    }
    merge_resource(parts, resource)
}
fn add_relationship_plan(
    plans: &mut BTreeMap<String, (String, String)>,
    id: &str,
    kind: &str,
    target: &str,
) -> Result<()> {
    validate_id(id)?;
    let plan = (kind.to_owned(), target.to_owned());
    if let Some(old) = plans.get(id) {
        if old != &plan {
            return Err(invalid(format!("conflicting relationship ID '{id}'")));
        }
    } else {
        plans.insert(id.to_owned(), plan);
    }
    Ok(())
}
fn add_payload(total: &mut usize, size: usize) -> Result<()> {
    if size > MAX_PAYLOAD_BYTES {
        return Err(limit("individual payload bytes"));
    }
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("total payload bytes"))?;
    if *total > MAX_TOTAL_PAYLOAD_BYTES {
        Err(limit("total payload bytes"))
    } else {
        Ok(())
    }
}
fn is_media_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::AUDIO | rt::VIDEO | rt::MEDIA | STRICT_AUDIO_REL | STRICT_VIDEO_REL
    )
}
fn is_media_content_type(value: &str, kind: SlideMediaKind) -> bool {
    match kind {
        SlideMediaKind::Audio => value.starts_with("audio/"),
        SlideMediaKind::Video => value.starts_with("video/") || value == "application/vnd.ms-asf",
    }
}
fn is_image_content_type(value: &str) -> bool {
    value.starts_with("image/") || matches!(value, "application/x-emf" | "application/x-wmf")
}
fn require_slide(part: &dyn Part) -> Result<()> {
    if part.content_type() == ct::PML_SLIDE {
        Ok(())
    } else {
        Err(invalid(format!(
            "part '{}' is not a slide",
            part.partname()
        )))
    }
}

fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("input XML bytes"));
    }
    let mut caps = MceCapabilities::ooxml_baseline();
    caps.understand_namespace(P14);
    let limits = MceLimits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &caps, &limits)?.xml;
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let node = make_node(&reader, element, reader.decoder(), &mut strings)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside slide root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|v| v.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside slide root"));
                }
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected in slide media markup")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated slide XML"));
    }
    root.ok_or_else(|| invalid("missing slide root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + name.len())?;
    let mut attributes = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let qname = item.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if attributes
            .iter()
            .any(|attribute: &Attribute| attribute.namespace == namespace && attribute.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}
fn conformance(root: &Node) -> Result<SlideMediaConformance> {
    crate_conformance(root)
}
fn crate_conformance(root: &Node) -> Result<SlideMediaConformance> {
    if root.name != "sld" {
        return Err(invalid("expected Slide root"));
    }
    match root.namespace.as_str() {
        PML => Ok(SlideMediaConformance::Transitional),
        STRICT_PML => Ok(SlideMediaConformance::Strict),
        _ => Err(invalid("unsupported Slide namespace")),
    }
}
fn one_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<Option<&'a Node>> {
    let mut values = node
        .children
        .iter()
        .filter(|child| child.namespace == namespace && child.name == name);
    let value = values.next();
    if values.next().is_some() {
        Err(invalid(format!(
            "{} has multiple {name} children",
            node.name
        )))
    } else {
        Ok(value)
    }
}
fn required_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a Node> {
    one_child(node, namespace, name)?
        .ok_or_else(|| invalid(format!("{} is missing {name}", node.name)))
}
fn require_node(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!("expected {name}, got {}", node.name)))
    }
}
fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}
fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name)))
}
fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    if let Some(attribute) = node.attributes.iter().find(|attribute| {
        !allowed.contains(&(attribute.namespace.as_str(), attribute.name.as_str()))
    }) {
        Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )))
    } else {
        Ok(())
    }
}
fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}
fn leaf(node: &Node, name: &str) -> Result<()> {
    whitespace(node)?;
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{name} must not contain child elements")))
    }
}
fn parse_i64(value: &str, name: &str) -> Result<i64> {
    value
        .parse()
        .map_err(|_| invalid(format!("invalid media transform {name}")))
}
fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID cannot be empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!("invalid relationship ID '{value}'")))
    } else {
        Ok(())
    }
}
fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("string bytes"))
    }
}
fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => Ok(std::str::from_utf8(value.as_ref())
            .map_err(xml_error)?
            .to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
fn attr(output: &mut Vec<u8>, name: &str, value: &str) {
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
    invalid(format!("slide media {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI_AUDIO: &[u8] =
        include_bytes!("../../../../test-data/poi/test-data/slideshow/EmbeddedAudio.pptx");
    const POI_VIDEO: &[u8] =
        include_bytes!("../../../../test-data/poi/test-data/slideshow/EmbeddedVideo.pptx");

    fn extension() -> OfficeMediaExtension {
        OfficeMediaExtension {
            embed_relationship_id: Some("rIdMedia".into()),
            link_relationship_id: None,
            trim: Some(MediaTrim {
                start: Some("1.5s".into()),
                end: Some("250ms".into()),
            }),
            fade: Some(MediaFade {
                fade_in: Some("1000".into()),
                fade_out: Some("2s".into()),
            }),
            bookmarks: vec![MediaBookmark {
                name: Some("chapter".into()),
                time: Some("3s".into()),
            }],
        }
    }
    fn value() -> SlideMediaList {
        SlideMediaList {
            pictures: vec![SlideMediaPicture {
                shape_id: 4,
                name: "sample.mp4".into(),
                kind: SlideMediaKind::Video,
                relationship_id: "rIdVideo".into(),
                resource: Some(MediaResource {
                    part_name: "/ppt/media/media1.mp4".into(),
                    content_type: "video/mp4".into(),
                    data: vec![0, 1, 2, 3],
                }),
                poster: Some(SlideMediaPoster {
                    relationship_id: "rIdPoster".into(),
                    resource: Some(MediaResource {
                        part_name: "/ppt/media/image1.png".into(),
                        content_type: "image/png".into(),
                        data: vec![137, 80, 78, 71],
                    }),
                }),
                transform: Some(SlideMediaTransform {
                    x: 100,
                    y: 200,
                    width: 300,
                    height: 400,
                }),
                office_extension: Some(extension()),
            }],
        }
    }
    fn slide_xml(conformance: SlideMediaConformance, inside: &[u8]) -> Vec<u8> {
        [format!("<p:sld xmlns:p=\"{}\" xmlns:a=\"{}\" xmlns:r=\"{}\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>", conformance.pml(), conformance.dml(), conformance.rel()).as_bytes(), inside, b"</p:spTree></p:cSld></p:sld>"].concat()
    }
    fn package(conformance: SlideMediaConformance) -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            ct::PML_SLIDE.into(),
            slide_xml(conformance, b""),
        )));
        (package, uri)
    }

    #[test]
    fn strict_xml_round_trip_covers_typed_extension_properties() {
        let expected = value();
        let fragment =
            write_slide_media_pictures(&expected, SlideMediaConformance::Strict).unwrap();
        let parsed =
            parse_slide_media(&slide_xml(SlideMediaConformance::Strict, &fragment)).unwrap();
        assert_eq!(parsed.pictures[0].shape_id, 4);
        assert_eq!(parsed.pictures[0].transform.unwrap().width, 300);
        let extension = parsed.pictures[0].office_extension.as_ref().unwrap();
        assert_eq!(
            extension.trim.as_ref().unwrap().start.as_deref(),
            Some("1.5s")
        );
        assert_eq!(extension.bookmarks[0].name.as_deref(), Some("chapter"));
        assert!(parsed.pictures[0].resource.is_none());
    }

    #[test]
    fn mce_fallback_selects_supported_media_picture() {
        let fragment =
            write_slide_media_pictures(&value(), SlideMediaConformance::Transitional).unwrap();
        let alternate = [b"<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x=\"urn:unsupported\"><mc:Choice Requires=\"x\"><p:pic/></mc:Choice><mc:Fallback>".as_slice(), fragment.as_slice(), b"</mc:Fallback></mc:AlternateContent>"].concat();
        assert_eq!(
            parse_slide_media(&slide_xml(SlideMediaConformance::Transitional, &alternate))
                .unwrap()
                .pictures
                .len(),
            1
        );
    }

    #[test]
    fn loads_poi_audio_and_video_resources_without_decoding() {
        for (bytes, kind, size) in [
            (POI_AUDIO, SlideMediaKind::Audio, 52_079usize),
            (POI_VIDEO, SlideMediaKind::Video, 101_799usize),
        ] {
            let package = OpcPackage::from_bytes(bytes).unwrap();
            let uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
            let media = load_slide_media(&package, &uri).unwrap();
            assert_eq!(media.pictures.len(), 1);
            let picture = &media.pictures[0];
            assert_eq!(picture.kind, kind);
            assert_eq!(picture.resource.as_ref().unwrap().data.len(), size);
            assert!(
                picture
                    .poster
                    .as_ref()
                    .unwrap()
                    .resource
                    .as_ref()
                    .unwrap()
                    .content_type
                    .starts_with("image/")
            );
            assert!(
                picture
                    .office_extension
                    .as_ref()
                    .unwrap()
                    .embed_relationship_id
                    .is_some()
            );
        }
    }

    #[test]
    fn transitional_package_writer_round_trips_complete_graph() {
        let (mut package, uri) = package(SlideMediaConformance::Transitional);
        let expected = value();
        store_slide_media(
            &mut package,
            &uri,
            &expected,
            SlideMediaConformance::Transitional,
        )
        .unwrap();
        assert_eq!(load_slide_media(&package, &uri).unwrap(), expected);
    }

    #[test]
    fn rejects_malformed_markup_caps_and_package_graphs() {
        let malformed = format!(
            "<p:sld xmlns:p=\"{PML}\" xmlns:a=\"{DML}\" xmlns:r=\"{REL}\"><p:pic><p:nvPicPr><p:cNvPr id=\"1\"/><p:nvPr><a:audioFile r:link=\"rId1\"/><p:extLst><p:ext><p14:media xmlns:p14=\"{P14}\" r:embed=\"rId2\"><p14:trim st=\"1..2s\"/></p14:media></p:ext></p:extLst></p:nvPr></p:nvPicPr></p:pic></p:sld>"
        );
        assert!(parse_slide_media(malformed.as_bytes()).is_err());
        assert!(parse_slide_media(b"<!DOCTYPE x><p:sld/>").is_err());
        assert!(parse_slide_media(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
        let (mut package, uri) = package(SlideMediaConformance::Transitional);
        let mut expected = value();
        for picture in &mut expected.pictures {
            picture.resource = None;
            if let Some(poster) = picture.poster.as_mut() {
                poster.resource = None;
            }
        }
        let fragment =
            write_slide_media_pictures(&expected, SlideMediaConformance::Transitional).unwrap();
        package
            .get_part_mut(&uri)
            .unwrap()
            .set_blob(slide_xml(SlideMediaConformance::Transitional, &fragment));
        assert!(load_slide_media(&package, &uri).is_err());
    }
}
