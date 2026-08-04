//! Typed PresentationML audio/video pictures and inert package media resources.
//!
//! The media-part graph follows [MS-PPTX] §2.1.1 and the media extension
//! vocabulary in §2.2.4. Binary payloads remain inert and are shared by Arc.

use crate::time::{Offset, ParseError as TimeParseError};
use crate::{Error, Result};
use litchi_drawingml::coord::{Coordinate, Extent, ParseError as CoordinateParseError};
use litchi_ooxml_common::mce::MCE_NAMESPACE;
use litchi_ooxml_common::{ExpandedName, MceCapabilities, MceLimits, process_markup_compatibility};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet, HashSet};
use std::sync::Arc;

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
const MAX_MEDIA_EXTENSION_XML_BYTES: usize = 4 * 1024 * 1024;
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

/// Immutable media bytes with copy-free clones and slice-style access.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaData(Arc<Vec<u8>>);

impl MediaData {
    fn from_shared(data: Arc<Vec<u8>>) -> Self {
        Self(data)
    }

    fn into_shared(self) -> Arc<Vec<u8>> {
        self.0
    }

    /// Borrow the inert payload bytes.
    #[must_use]
    pub fn as_slice(&self) -> &[u8] {
        self.0.as_slice()
    }

    /// Return the payload length in bytes.
    #[must_use]
    pub fn len(&self) -> usize {
        self.0.len()
    }

    /// Return whether the payload is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    /// Recover the backing vector without copying when this is its sole owner.
    ///
    /// On sharing, ownership of `self` is returned so callers can keep borrowing
    /// the bytes or deliberately choose a copy.
    pub fn try_into_vec(self) -> std::result::Result<Vec<u8>, Self> {
        Arc::try_unwrap(self.0).map_err(Self)
    }

    /// Return whether two values share the same immutable allocation.
    #[must_use]
    pub fn shares_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }
}

impl From<Vec<u8>> for MediaData {
    fn from(value: Vec<u8>) -> Self {
        Self(Arc::new(value))
    }
}

impl AsRef<[u8]> for MediaData {
    fn as_ref(&self) -> &[u8] {
        self.as_slice()
    }
}

impl std::ops::Deref for MediaData {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        self.as_slice()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaResource {
    pub part_name: String,
    pub content_type: String,
    /// Stored and returned verbatim. The payload is never decoded or executed.
    /// Clones share these immutable bytes instead of copying large media parts.
    pub data: MediaData,
}

impl MediaResource {
    /// Construct an inert resource while moving its payload into shared storage.
    #[must_use]
    pub fn new(
        part_name: impl Into<String>,
        content_type: impl Into<String>,
        data: impl Into<MediaData>,
    ) -> Self {
        Self {
            part_name: part_name.into(),
            content_type: content_type.into(),
            data: data.into(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideMediaPoster {
    pub relationship_id: String,
    pub resource: Option<MediaResource>,
}

/// A checked DrawingML transform for a media picture.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideMediaTransform {
    x: Coordinate,
    y: Coordinate,
    width: Extent,
    height: Extent,
}

impl SlideMediaTransform {
    /// Construct a transform from schema-checked DrawingML values.
    #[must_use]
    pub fn new(x: Coordinate, y: Coordinate, width: Extent, height: Extent) -> Self {
        Self {
            x,
            y,
            width,
            height,
        }
    }

    /// Construct a transform from EMUs with all schema bounds checked.
    pub fn emu(x: i64, y: i64, width: i64, height: i64) -> Result<Self> {
        Ok(Self::new(
            Coordinate::emu(x).map_err(|error| coordinate_error(error, "x"))?,
            Coordinate::emu(y).map_err(|error| coordinate_error(error, "y"))?,
            Extent::emu(width).map_err(|error| coordinate_error(error, "width"))?,
            Extent::emu(height).map_err(|error| coordinate_error(error, "height"))?,
        ))
    }

    /// Borrow the horizontal offset.
    pub const fn x(&self) -> &Coordinate {
        &self.x
    }

    /// Borrow the vertical offset.
    pub const fn y(&self) -> &Coordinate {
        &self.y
    }

    /// Borrow the schema-checked horizontal extent.
    pub const fn width(&self) -> &Extent {
        &self.width
    }

    /// Borrow the schema-checked vertical extent.
    pub const fn height(&self) -> &Extent {
        &self.height
    }
}

/// Typed amounts removed from the beginning and end of media playback.
///
/// The media-length-dependent sum constraint is intentionally local to
/// media-aware validation; an [`Offset`] only validates its own exact value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MediaTrim {
    /// Authored `st`; absence has an effective value of zero.
    pub start: Option<Offset>,
    /// Authored `end`; absence has an effective value of zero.
    pub end: Option<Offset>,
}

impl MediaTrim {
    /// Borrow the effective start offset without erasing authored absence.
    pub fn start(&self) -> &Offset {
        self.start.as_ref().unwrap_or(&Offset::ZERO)
    }

    /// Borrow the effective end offset without erasing authored absence.
    pub fn end(&self) -> &Offset {
        self.end.as_ref().unwrap_or(&Offset::ZERO)
    }
}

/// Typed fade durations at the beginning and end of media playback.
///
/// The combined-duration constraint requires the media length and is therefore
/// kept out of the reusable [`Offset`] value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MediaFade {
    /// Authored `in`; absence has an effective value of zero.
    pub fade_in: Option<Offset>,
    /// Authored `out`; absence has an effective value of zero.
    pub fade_out: Option<Offset>,
}

impl MediaFade {
    /// Borrow the effective fade-in duration without erasing authored absence.
    pub fn fade_in(&self) -> &Offset {
        self.fade_in.as_ref().unwrap_or(&Offset::ZERO)
    }

    /// Borrow the effective fade-out duration without erasing authored absence.
    pub fn fade_out(&self) -> &Offset {
        self.fade_out.as_ref().unwrap_or(&Offset::ZERO)
    }
}

/// A named point on a media timeline.
///
/// Time uniqueness is semantic. The upper bound against the actual media
/// length remains a media-aware check because payloads are not decoded here.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MediaBookmark {
    pub name: Option<String>,
    pub time: Option<Offset>,
}

/// A bounded, canonical `p:extLst` fragment retained without interpretation.
///
/// The wrapper is validated and canonicalized while retaining QName prefixes
/// and their bindings. Extension payloads remain inert and are never loaded,
/// dispatched, or executed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MediaExtensionList {
    xml: Box<str>,
}

impl MediaExtensionList {
    /// Parse one transitional PresentationML `p:extLst` fragment.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_MEDIA_EXTENSION_XML_BYTES {
            return Err(limit("media extension-list XML bytes"));
        }
        let root = parse_document(xml)?;
        Self::from_node(&root)
    }

    /// Borrow the self-contained canonical XML fragment.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.xml
    }

    fn from_node(node: &Node) -> Result<Self> {
        require_node(node, PML, "extLst")?;
        let xml = canonical_fragment(node)?;
        let xml = String::from_utf8(xml)
            .map_err(|_| invalid("canonical media extension-list XML is not UTF-8"))?;
        Ok(Self {
            xml: xml.into_boxed_str(),
        })
    }
}

impl AsRef<str> for MediaExtensionList {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<&str> for MediaExtensionList {
    type Error = crate::Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::parse(value.as_bytes())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OfficeMediaExtension {
    pub embed_relationship_id: Option<String>,
    pub link_relationship_id: Option<String>,
    pub trim: Option<MediaTrim>,
    pub fade: Option<MediaFade>,
    pub bookmarks: Vec<MediaBookmark>,
    /// Optional opaque PresentationML extension metadata, ordered last by XSD.
    pub extensions: Option<MediaExtensionList>,
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
    prefix: String,
    name: String,
    value: String,
}

#[derive(Default)]
struct NamespaceContext {
    parent: Option<Arc<Self>>,
    declarations: Vec<(String, String)>,
}

#[derive(Clone)]
struct Node {
    namespace: String,
    prefix: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    /// Direct text plus UTF-8 boundaries immediately before each child.
    text: String,
    text_ends: Vec<usize>,
    namespace_context: Arc<NamespaceContext>,
    declares_namespaces: bool,
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
    if node.namespace == conformance.pml()
        && node.name == "pic"
        && let Some(picture) = parse_picture(node, conformance)?
    {
        if output.len() == MAX_MEDIA {
            return Err(limit("media count"));
        }
        output.push(picture);
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
        .map(parse_office_media)
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
    Ok(Some(SlideMediaTransform::new(
        parse_coordinate(required(offset, "", "x")?, "x")?,
        parse_coordinate(required(offset, "", "y")?, "y")?,
        parse_extent(required(extent, "", "cx")?, "width")?,
        parse_extent(required(extent, "", "cy")?, "height")?,
    )))
}

fn find_office_media(node: &Node, conformance: SlideMediaConformance) -> Result<Option<&Node>> {
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

fn parse_office_media(node: &Node) -> Result<OfficeMediaExtension> {
    whitespace(node)?;
    // The Office 2010 p14 schema imports the transitional relationships
    // namespace even when the containing presentation uses Strict namespaces.
    let embed_relationship_id = optional(node, REL, "embed").map(str::to_owned);
    let link_relationship_id = optional(node, REL, "link").map(str::to_owned);
    no_attributes(node, &[(REL, "embed"), (REL, "link")])?;
    if embed_relationship_id.is_none() && link_relationship_id.is_none() {
        return Err(invalid("p14:media requires r:embed or r:link"));
    }
    let mut trim_node = None;
    let mut fade_node = None;
    let mut bookmarks_node = None;
    let mut extensions = None;
    let mut stage = 0u8;
    for child in &node.children {
        let child_stage = match (child.namespace.as_str(), child.name.as_str()) {
            (P14, "trim") => {
                if trim_node.replace(child).is_some() {
                    return Err(invalid("p14:media has multiple trim children"));
                }
                1
            },
            (P14, "fade") => {
                if fade_node.replace(child).is_some() {
                    return Err(invalid("p14:media has multiple fade children"));
                }
                2
            },
            (P14, "bmkLst") => {
                if bookmarks_node.replace(child).is_some() {
                    return Err(invalid("p14:media has multiple bmkLst children"));
                }
                3
            },
            // The p14 schema imports transitional PresentationML even when
            // the containing slide uses the Strict namespace dialect.
            (PML, "extLst") => {
                if extensions
                    .replace(MediaExtensionList::from_node(child)?)
                    .is_some()
                {
                    return Err(invalid("p14:media has multiple extLst children"));
                }
                4
            },
            _ => {
                return Err(invalid(format!(
                    "unsupported p14:media child '{}'",
                    child.name
                )));
            },
        };
        if child_stage < stage {
            return Err(invalid(
                "p14:media children are outside the schema-defined order",
            ));
        }
        stage = child_stage;
    }
    let trim = trim_node.map(parse_trim).transpose()?;
    let fade = fade_node.map(parse_fade).transpose()?;
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
                time: optional(child, "", "time").map(parse_time).transpose()?,
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
        extensions,
    })
}

fn parse_trim(node: &Node) -> Result<MediaTrim> {
    leaf(node, "media trim")?;
    no_attributes(node, &[("", "st"), ("", "end")])?;
    Ok(MediaTrim {
        start: optional(node, "", "st").map(parse_time).transpose()?,
        end: optional(node, "", "end").map(parse_time).transpose()?,
    })
}

fn parse_fade(node: &Node) -> Result<MediaFade> {
    leaf(node, "media fade")?;
    no_attributes(node, &[("", "in"), ("", "out")])?;
    Ok(MediaFade {
        fade_in: optional(node, "", "in").map(parse_time).transpose()?,
        fade_out: optional(node, "", "out").map(parse_time).transpose()?,
    })
}

/// Deterministically serializes self-contained `p:pic` fragments.
pub fn write_slide_media_pictures(
    value: &SlideMediaList,
    conformance: SlideMediaConformance,
) -> Result<Vec<u8>> {
    validate_value(value, false)?;
    let mut output = BoundedXml::new();
    for picture in &value.pictures {
        write_picture(&mut output, picture, conformance)?;
    }
    Ok(output.finish())
}

fn write_picture(
    output: &mut BoundedXml,
    picture: &SlideMediaPicture,
    conformance: SlideMediaConformance,
) -> Result<()> {
    output.write(b"<p:pic xmlns:p=\"")?;
    output.escape(conformance.pml())?;
    output.write(b"\" xmlns:a=\"")?;
    output.escape(conformance.dml())?;
    output.write(b"\" xmlns:r=\"")?;
    output.escape(conformance.rel())?;
    if picture.office_extension.is_some() {
        output.write(b"\" xmlns:p14=\"")?;
        output.escape(P14)?;
    }
    output.write(b"\"><p:nvPicPr><p:cNvPr")?;
    output.attr("id", &picture.shape_id.to_string())?;
    output.attr("name", &picture.name)?;
    output.write(b"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr>")?;
    output.write(match picture.kind {
        SlideMediaKind::Audio => b"<a:audioFile".as_slice(),
        SlideMediaKind::Video => b"<a:videoFile".as_slice(),
    })?;
    output.attr("r:link", &picture.relationship_id)?;
    output.write(b"/>")?;
    if let Some(extension) = &picture.office_extension {
        write_office_extension(output, extension)?;
    }
    output.write(b"</p:nvPr></p:nvPicPr><p:blipFill>")?;
    if let Some(poster) = &picture.poster {
        output.write(b"<a:blip")?;
        output.attr("r:embed", &poster.relationship_id)?;
        output.write(b"/><a:stretch><a:fillRect/></a:stretch>")?;
    }
    output.write(b"</p:blipFill><p:spPr>")?;
    if let Some(transform) = &picture.transform {
        output.write(b"<a:xfrm><a:off")?;
        output.attr("x", &transform.x.to_string())?;
        output.attr("y", &transform.y.to_string())?;
        output.write(b"/><a:ext")?;
        output.attr("cx", &transform.width.to_string())?;
        output.attr("cy", &transform.height.to_string())?;
        output.write(b"/></a:xfrm>")?;
    }
    output.write(b"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>")?;
    Ok(())
}

fn write_office_extension(output: &mut BoundedXml, value: &OfficeMediaExtension) -> Result<()> {
    output.write(b"<p:extLst><p:ext uri=\"")?;
    output.escape(MEDIA_EXTENSION_URI)?;
    output.write(b"\"><p14:media xmlns:r=\"")?;
    output.escape(REL)?;
    output.push(b'"')?;
    if let Some(id) = &value.embed_relationship_id {
        output.attr("r:embed", id)?;
    }
    if let Some(id) = &value.link_relationship_id {
        output.attr("r:link", id)?;
    }
    if value.trim.is_none()
        && value.fade.is_none()
        && value.bookmarks.is_empty()
        && value.extensions.is_none()
    {
        output.write(b"/></p:ext></p:extLst>")?;
        return Ok(());
    }
    output.push(b'>')?;
    if let Some(trim) = &value.trim {
        output.write(b"<p14:trim")?;
        if let Some(start) = &trim.start {
            output.attr("st", start.as_str())?;
        }
        if let Some(end) = &trim.end {
            output.attr("end", end.as_str())?;
        }
        output.write(b"/>")?;
    }
    if let Some(fade) = &value.fade {
        output.write(b"<p14:fade")?;
        if let Some(fade_in) = &fade.fade_in {
            output.attr("in", fade_in.as_str())?;
        }
        if let Some(fade_out) = &fade.fade_out {
            output.attr("out", fade_out.as_str())?;
        }
        output.write(b"/>")?;
    }
    if !value.bookmarks.is_empty() {
        output.write(b"<p14:bmkLst>")?;
        for bookmark in &value.bookmarks {
            output.write(b"<p14:bmk")?;
            if let Some(v) = &bookmark.name {
                output.attr("name", v)?;
            }
            if let Some(v) = &bookmark.time {
                output.attr("time", v.as_str())?;
            }
            output.write(b"/>")?;
        }
        output.write(b"</p14:bmkLst>")?;
    }
    if let Some(extensions) = &value.extensions {
        output.write(extensions.as_str().as_bytes())?;
    }
    output.write(b"</p14:media></p:ext></p:extLst>")?;
    Ok(())
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
    relationship.target_partname().map_err(Error::Opc)
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
        data: MediaData::from_shared(part.blob_arc()),
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
        let uri = PackURI::new(&resource.part_name).map_err(Error::Invalid)?;
        package.add_part(Box::new(BlobPart::new_shared(
            uri,
            resource.content_type,
            resource.data.into_shared(),
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
                let core = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == conformance.pml().as_bytes());
                if depth == 0 && (!core || element.local_name().as_ref() != b"sld") {
                    return Err(invalid("slide root does not match conformance"));
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return Err(limit("XML depth"));
                }
                if core
                    && element.local_name().as_ref() == b"spTree"
                    && sp_tree_depth.replace(depth).is_some()
                {
                    return Err(invalid("slide has multiple shape trees"));
                }
            },
            Event::Empty(element) if element.local_name().as_ref() == b"spTree" => {
                return Err(invalid("cannot insert into an empty shape tree"));
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
    let mut payload_allocations = HashSet::new();
    let mut resident_payload = 0usize;
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
        if require_resources && picture.resource.is_none() {
            return Err(invalid("media resource is required for package storage"));
        }
        if let Some(resource) = &picture.resource {
            resource_uri(resource, false, Some(picture.kind))?;
            count_payload_allocation(
                &mut payload_allocations,
                &mut resident_payload,
                &resource.data,
            )?;
            merge_resource(&mut resources, resource)?;
        }
        if let Some(poster) = &picture.poster {
            validate_id(&poster.relationship_id)?;
            if require_resources && poster.resource.is_none() {
                return Err(invalid("poster resource is required for package storage"));
            }
            if let Some(resource) = &poster.resource {
                resource_uri(resource, true, None)?;
                count_payload_allocation(
                    &mut payload_allocations,
                    &mut resident_payload,
                    &resource.data,
                )?;
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

fn count_payload_allocation(
    allocations: &mut HashSet<(usize, usize)>,
    total: &mut usize,
    data: &MediaData,
) -> Result<()> {
    let allocation = (data.as_ptr() as usize, data.len());
    if allocations.insert(allocation) {
        add_payload(total, data.len())?;
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
        if let Some(time) = &bookmark.time
            && !times.insert(time)
        {
            return Err(invalid(format!("duplicate media bookmark time '{time}'")));
        }
    }
    Ok(())
}

fn parse_time(value: &str) -> Result<Offset> {
    Offset::parse(value).map_err(time_error)
}

fn time_error(error: TimeParseError) -> Error {
    invalid(format!("invalid media universal time offset: {error}"))
}

fn resource_uri(
    resource: &MediaResource,
    image: bool,
    kind: Option<SlideMediaKind>,
) -> Result<PackURI> {
    let uri = PackURI::new(&resource.part_name).map_err(Error::Invalid)?;
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
    } else {
        let kind = kind.ok_or_else(|| invalid("non-image media resource requires a media kind"))?;
        if !is_media_content_type(&resource.content_type, kind) {
            return Err(invalid(
                "media resource content type is inconsistent with its kind",
            ));
        }
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
    for namespace in [PML, STRICT_PML] {
        caps.preserve_extension_element(ExpandedName {
            namespace: namespace.to_owned(),
            local_name: "ext".to_owned(),
        });
    }
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
    let base_namespace_context = Arc::new(NamespaceContext::default());
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let parent_namespace_context = stack
                    .last()
                    .map_or(&base_namespace_context, |node| &node.namespace_context);
                let node = make_node(
                    &reader,
                    element,
                    reader.decoder(),
                    &mut strings,
                    parent_namespace_context,
                )?;
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
            Event::CData(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("CDATA outside slide root"));
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
        return Err(invalid("unterminated slide XML"));
    }
    root.ok_or_else(|| invalid("missing slide root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
    parent_namespace_context: &Arc<NamespaceContext>,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let prefix = element
        .name()
        .prefix()
        .map(|prefix| std::str::from_utf8(prefix.as_ref()).map(str::to_owned))
        .transpose()
        .map_err(xml_error)?
        .unwrap_or_default();
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + prefix.len() + name.len())?;
    let mut attributes = Vec::new();
    let mut namespace_declarations = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if let Some(declaration) = item.key.as_namespace_binding() {
            let prefix = match declaration {
                PrefixDeclaration::Default => String::new(),
                PrefixDeclaration::Named(prefix) => {
                    std::str::from_utf8(prefix).map_err(xml_error)?.to_owned()
                },
            };
            add_strings(strings, prefix.len() + value.len())?;
            namespace_declarations.push((prefix, value));
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let prefix = item
            .key
            .prefix()
            .map(|prefix| std::str::from_utf8(prefix.as_ref()).map(str::to_owned))
            .transpose()
            .map_err(xml_error)?
            .unwrap_or_default();
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        add_strings(
            strings,
            namespace.len() + prefix.len() + name.len() + value.len(),
        )?;
        if attributes
            .iter()
            .any(|attribute: &Attribute| attribute.namespace == namespace && attribute.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            prefix,
            name,
            value,
        });
    }
    let declares_namespaces = !namespace_declarations.is_empty();
    let namespace_context = if declares_namespaces {
        Arc::new(NamespaceContext {
            parent: Some(Arc::clone(parent_namespace_context)),
            declarations: namespace_declarations,
        })
    } else {
        Arc::clone(parent_namespace_context)
    };
    Ok(Node {
        namespace,
        prefix,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
        text_ends: Vec::new(),
        namespace_context,
        declares_namespaces,
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.text_ends.push(parent.text.len());
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
fn parse_coordinate(value: &str, name: &str) -> Result<Coordinate> {
    Coordinate::parse(value).map_err(|error| coordinate_error(error, name))
}

fn parse_extent(value: &str, name: &str) -> Result<Extent> {
    Extent::parse(value).map_err(|error| coordinate_error(error, name))
}

fn coordinate_error(error: CoordinateParseError, name: &str) -> Error {
    invalid(format!("invalid media transform {name}: {error}"))
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
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn canonical_fragment(node: &Node) -> Result<Vec<u8>> {
    let mut used_prefixes = BTreeSet::new();
    collect_used_prefixes(node, &mut used_prefixes);
    let mut output = BoundedXml::with_limit(MAX_MEDIA_EXTENSION_XML_BYTES);
    write_canonical_node(&mut output, node, &used_prefixes, true)?;
    Ok(output.finish())
}

fn collect_used_prefixes(node: &Node, prefixes: &mut BTreeSet<String>) {
    if !node.prefix.is_empty() {
        prefixes.insert(node.prefix.clone());
    } else if !node.namespace.is_empty() {
        prefixes.insert(String::new());
    }
    for attribute in &node.attributes {
        if !attribute.prefix.is_empty() {
            prefixes.insert(attribute.prefix.clone());
        }
        for token in attribute.value.split(|character: char| {
            character.is_whitespace()
                || matches!(
                    character,
                    ',' | ';' | '(' | ')' | '[' | ']' | '{' | '}' | '/' | '\\' | '"' | '\''
                )
        }) {
            if let Some((prefix, _)) = token.split_once(':')
                && !prefix.is_empty()
            {
                prefixes.insert(prefix.to_owned());
            }
        }
        if attribute.namespace == MCE_NAMESPACE
            && matches!(attribute.name.as_str(), "Ignorable" | "Requires")
        {
            prefixes.extend(
                attribute
                    .value
                    .split_whitespace()
                    .filter(|prefix| !prefix.is_empty())
                    .map(str::to_owned),
            );
        }
    }
    for child in &node.children {
        collect_used_prefixes(child, prefixes);
    }
}

fn write_canonical_node(
    output: &mut BoundedXml,
    node: &Node,
    used_prefixes: &BTreeSet<String>,
    root: bool,
) -> Result<()> {
    if node.text_ends.len() != node.children.len() {
        return Err(invalid("invalid XML text-segment state"));
    }

    output.push(b'<')?;
    write_original_name(output, &node.prefix, &node.name)?;
    if root {
        let mut bindings = effective_namespace_bindings(&node.namespace_context)?;
        bindings.retain(|prefix, _| used_prefixes.contains(prefix));
        bindings.entry(String::new()).or_default();
        for (prefix, namespace) in &bindings {
            if prefix != "xml" {
                write_namespace_declaration(output, prefix, namespace)?;
            }
        }
    } else if node.declares_namespaces {
        let mut declarations: Vec<_> = node.namespace_context.declarations.iter().collect();
        declarations.sort_unstable_by(|left, right| left.0.cmp(&right.0));
        for (prefix, namespace) in declarations {
            if used_prefixes.contains(prefix) {
                write_namespace_declaration(output, prefix, namespace)?;
            }
        }
    }

    let mut attributes: Vec<_> = node.attributes.iter().collect();
    attributes.sort_unstable_by(|left, right| {
        (&left.namespace, &left.name, &left.prefix).cmp(&(
            &right.namespace,
            &right.name,
            &right.prefix,
        ))
    });
    for attribute in attributes {
        output.push(b' ')?;
        write_original_name(output, &attribute.prefix, &attribute.name)?;
        output.write(b"=\"")?;
        output.escape(&attribute.value)?;
        output.push(b'"')?;
    }

    if node.children.is_empty() && node.text.is_empty() {
        output.write(b"/>")?;
        return Ok(());
    }
    output.push(b'>')?;
    let mut text_start = 0usize;
    for (index, child) in node.children.iter().enumerate() {
        let text_end = *node
            .text_ends
            .get(index)
            .ok_or_else(|| invalid("invalid XML text-segment state"))?;
        let text = node
            .text
            .get(text_start..text_end)
            .ok_or_else(|| invalid("invalid XML text-segment state"))?;
        output.escape(text)?;
        write_canonical_node(output, child, used_prefixes, false)?;
        text_start = text_end;
    }
    let text = node
        .text
        .get(text_start..)
        .ok_or_else(|| invalid("invalid XML text-segment state"))?;
    output.escape(text)?;
    output.write(b"</")?;
    write_original_name(output, &node.prefix, &node.name)?;
    output.push(b'>')?;
    Ok(())
}

fn effective_namespace_bindings(
    context: &Arc<NamespaceContext>,
) -> Result<BTreeMap<String, String>> {
    let mut contexts = Vec::new();
    let mut current = Some(context.as_ref());
    while let Some(context) = current {
        if contexts.len() == MAX_DEPTH {
            return Err(limit("media extension namespace depth"));
        }
        contexts.push(context);
        current = context.parent.as_deref();
    }

    let mut bindings = BTreeMap::new();
    for context in contexts.into_iter().rev() {
        for (prefix, namespace) in &context.declarations {
            if namespace.is_empty() {
                bindings.remove(prefix);
            } else {
                bindings.insert(prefix.clone(), namespace.clone());
            }
        }
    }
    Ok(bindings)
}

fn write_namespace_declaration(
    output: &mut BoundedXml,
    prefix: &str,
    namespace: &str,
) -> Result<()> {
    if prefix.is_empty() {
        output.write(b" xmlns=\"")?;
    } else {
        output.write(b" xmlns:")?;
        output.write(prefix.as_bytes())?;
        output.write(b"=\"")?;
    }
    output.escape(namespace)?;
    output.push(b'"')
}

fn write_original_name(output: &mut BoundedXml, prefix: &str, name: &str) -> Result<()> {
    if !prefix.is_empty() {
        output.write(prefix.as_bytes())?;
        output.push(b':')?;
    }
    output.write(name.as_bytes())
}

struct BoundedXml {
    bytes: Vec<u8>,
    limit: usize,
}

impl BoundedXml {
    fn new() -> Self {
        Self::with_limit(MAX_XML_BYTES)
    }

    fn with_limit(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }

    fn write(&mut self, value: &[u8]) -> Result<()> {
        self.reserve(value.len())?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn push(&mut self, value: u8) -> Result<()> {
        self.reserve(1)?;
        self.bytes.push(value);
        Ok(())
    }

    fn escape(&mut self, value: &str) -> Result<()> {
        let escaped = escaped_len(value).ok_or_else(|| output_limit(self.limit))?;
        self.reserve(escaped)?;
        write_escaped(&mut self.bytes, value);
        Ok(())
    }

    fn attr(&mut self, name: &str, value: &str) -> Result<()> {
        let escaped = escaped_len(value).ok_or_else(|| output_limit(self.limit))?;
        let fixed = 4usize
            .checked_add(name.len())
            .ok_or_else(|| output_limit(self.limit))?;
        let required = fixed
            .checked_add(escaped)
            .ok_or_else(|| output_limit(self.limit))?;
        self.reserve(required)?;
        self.bytes.push(b' ');
        self.bytes.extend_from_slice(name.as_bytes());
        self.bytes.extend_from_slice(b"=\"");
        write_escaped(&mut self.bytes, value);
        self.bytes.push(b'\"');
        Ok(())
    }

    fn reserve(&mut self, additional: usize) -> Result<()> {
        let next = self
            .bytes
            .len()
            .checked_add(additional)
            .ok_or_else(|| output_limit(self.limit))?;
        if next > self.limit {
            return Err(output_limit(self.limit));
        }
        let spare = self.bytes.capacity().saturating_sub(self.bytes.len());
        if additional > spare {
            self.bytes
                .try_reserve_exact(additional)
                .map_err(|source| Error::Allocation {
                    resource: "slide media serialized XML",
                    source,
                })?;
        }
        Ok(())
    }
}

fn escaped_len(value: &str) -> Option<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let bytes = match character {
            '&' => 5,
            '<' => 4,
            '>' => 4,
            '"' => 6,
            '\t' => 5,
            '\n' | '\r' => 6,
            _ => character.len_utf8(),
        };
        total.checked_add(bytes)
    })
}

fn write_escaped(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
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
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn limit(name: &str) -> Error {
    invalid(format!("slide media {name} limit exceeded"))
}

fn output_limit(maximum: usize) -> Error {
    Error::Limit {
        resource: "slide media serialized XML bytes",
        limit: maximum,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI_AUDIO: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/slideshow/EmbeddedAudio.pptx");
    const POI_VIDEO: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/slideshow/EmbeddedVideo.pptx");

    fn extension() -> OfficeMediaExtension {
        OfficeMediaExtension {
            embed_relationship_id: Some("rIdMedia".into()),
            link_relationship_id: None,
            trim: Some(MediaTrim {
                start: Some(Offset::parse("1.5s").unwrap()),
                end: Some(Offset::ms(250)),
            }),
            fade: Some(MediaFade {
                fade_in: Some(Offset::ms(1000)),
                fade_out: Some(Offset::secs(2)),
            }),
            bookmarks: vec![MediaBookmark {
                name: Some("chapter".into()),
                time: Some(Offset::secs(3)),
            }],
            extensions: None,
        }
    }
    fn value() -> SlideMediaList {
        SlideMediaList {
            pictures: vec![SlideMediaPicture {
                shape_id: 4,
                name: "sample.mp4".into(),
                kind: SlideMediaKind::Video,
                relationship_id: "rIdVideo".into(),
                resource: Some(MediaResource::new(
                    "/ppt/media/media1.mp4",
                    "video/mp4",
                    vec![0, 1, 2, 3],
                )),
                poster: Some(SlideMediaPoster {
                    relationship_id: "rIdPoster".into(),
                    resource: Some(MediaResource::new(
                        "/ppt/media/image1.png",
                        "image/png",
                        vec![137, 80, 78, 71],
                    )),
                }),
                transform: Some(SlideMediaTransform::emu(100, 200, 300, 400).unwrap()),
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
        let xml = std::str::from_utf8(&fragment).unwrap();
        assert!(xml.contains(&format!(r#"xmlns:r="{STRICT_REL}""#)));
        assert!(xml.contains(&format!(r#"<p14:media xmlns:r="{REL}" r:embed="rIdMedia""#)));
        let parsed =
            parse_slide_media(&slide_xml(SlideMediaConformance::Strict, &fragment)).unwrap();
        assert_eq!(parsed.pictures[0].shape_id, 4);
        assert_eq!(
            parsed.pictures[0]
                .transform
                .as_ref()
                .unwrap()
                .width()
                .as_emu(),
            300
        );
        let extension = parsed.pictures[0].office_extension.as_ref().unwrap();
        assert_eq!(
            extension.trim.as_ref().unwrap().start,
            Some(Offset::ms(1500))
        );
        assert_eq!(extension.bookmarks[0].name.as_deref(), Some("chapter"));
        assert!(parsed.pictures[0].resource.is_none());
    }

    #[test]
    fn media_transform_round_trips_coordinate_offsets_and_integer_extents() {
        let mut expected = value();
        expected.pictures[0].transform = Some(SlideMediaTransform::new(
            Coordinate::parse("-1.25cm").unwrap(),
            Coordinate::emu(litchi_drawingml::coord::MIN_EMU).unwrap(),
            Extent::ZERO,
            Extent::emu(litchi_drawingml::coord::MAX_EMU).unwrap(),
        ));

        let fragment =
            write_slide_media_pictures(&expected, SlideMediaConformance::Transitional).unwrap();
        let xml = std::str::from_utf8(&fragment).unwrap();
        assert!(xml.contains(r#"<a:off x="-1.25cm" y="-27273042329600"/>"#));
        assert!(xml.contains(r#"<a:ext cx="0" cy="27273042316900"/>"#));

        let parsed =
            parse_slide_media(&slide_xml(SlideMediaConformance::Transitional, &fragment)).unwrap();
        assert_eq!(parsed.pictures[0].transform, expected.pictures[0].transform);
    }

    #[test]
    fn media_transform_construction_and_parsing_enforce_boundaries() {
        assert!(
            SlideMediaTransform::emu(
                litchi_drawingml::coord::MIN_EMU,
                litchi_drawingml::coord::MAX_EMU,
                1,
                litchi_drawingml::coord::MAX_EMU,
            )
            .is_ok()
        );
        assert!(SlideMediaTransform::emu(litchi_drawingml::coord::MIN_EMU - 1, 0, 1, 1).is_err());
        assert!(SlideMediaTransform::emu(0, 0, 0, 1).is_ok());
        assert!(SlideMediaTransform::emu(0, 0, 1, -1).is_err());
        assert!(SlideMediaTransform::emu(0, 0, litchi_drawingml::coord::MAX_EMU + 1, 1).is_err());

        let fragment =
            write_slide_media_pictures(&value(), SlideMediaConformance::Transitional).unwrap();
        let invalid = String::from_utf8(fragment)
            .unwrap()
            .replace(r#"cx="300""#, r#"cx="0mm""#);
        assert!(
            parse_slide_media(&slide_xml(
                SlideMediaConformance::Transitional,
                invalid.as_bytes(),
            ))
            .is_err()
        );
    }

    #[test]
    fn trim_and_fade_preserve_absence_and_explicit_zero() {
        let mut expected = value();
        let (authored_trim, authored_fade) = {
            let extension = expected.pictures[0].office_extension.as_mut().unwrap();
            extension.trim = Some(MediaTrim {
                start: Some(Offset::ZERO),
                end: None,
            });
            extension.fade = Some(MediaFade {
                fade_in: None,
                fade_out: Some(Offset::ZERO),
            });
            (extension.trim.clone(), extension.fade.clone())
        };

        let fragment =
            write_slide_media_pictures(&expected, SlideMediaConformance::Transitional).unwrap();
        let xml = std::str::from_utf8(&fragment).unwrap();
        assert!(xml.contains(r#"<p14:trim st="0"/>"#));
        assert!(xml.contains(r#"<p14:fade out="0"/>"#));

        let parsed =
            parse_slide_media(&slide_xml(SlideMediaConformance::Transitional, &fragment)).unwrap();
        let actual = parsed.pictures[0].office_extension.as_ref().unwrap();
        assert_eq!(actual.trim, authored_trim);
        assert_eq!(actual.fade, authored_fade);
        assert!(actual.trim.as_ref().unwrap().start().is_zero());
        assert!(actual.trim.as_ref().unwrap().end().is_zero());
        assert!(actual.fade.as_ref().unwrap().fade_in().is_zero());
        assert!(actual.fade.as_ref().unwrap().fade_out().is_zero());
    }

    #[test]
    fn opaque_media_extensions_round_trip_canonically() {
        let opaque = MediaExtensionList::parse(
            format!(
                r#"<p:extLst xmlns:p="{PML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:z="urn:example" mc:Ignorable="z"><p:ext uri="{{EXAMPLE}}"><z:data z:flag="a&amp;b">before<![CDATA[<literal>]]><z:inner/>after &amp; done</z:data></p:ext></p:extLst>"#
            )
            .as_bytes(),
        )
        .unwrap();
        assert!(opaque.as_str().contains("before&lt;literal&gt;<z:inner"));
        assert!(opaque.as_str().contains("after &amp; done"));

        let mut expected = value();
        expected.pictures[0]
            .office_extension
            .as_mut()
            .unwrap()
            .extensions = Some(opaque.clone());
        let fragment =
            write_slide_media_pictures(&expected, SlideMediaConformance::Strict).unwrap();
        let xml = std::str::from_utf8(&fragment).unwrap();
        assert!(xml.contains(&format!(r#"xmlns:p="{PML}""#)));
        assert!(xml.contains(r#"xmlns:z="urn:example""#));

        let parsed =
            parse_slide_media(&slide_xml(SlideMediaConformance::Strict, &fragment)).unwrap();
        assert_eq!(
            parsed.pictures[0]
                .office_extension
                .as_ref()
                .unwrap()
                .extensions
                .as_ref(),
            Some(&opaque)
        );
    }

    #[test]
    fn rejects_duplicate_and_misordered_media_children() {
        fn picture(children: &str) -> Vec<u8> {
            format!(
                r#"<p:pic xmlns:p14="{P14}"><p:nvPicPr><p:cNvPr id="1"/><p:nvPr><a:audioFile r:link="rId1"/><p:extLst><p:ext><p14:media xmlns:r="{REL}" r:embed="rId2">{children}</p14:media></p:ext></p:extLst></p:nvPr></p:nvPicPr></p:pic>"#
            )
            .into_bytes()
        }

        for children in [
            "<p14:trim/><p14:trim/>",
            "<p14:fade/><p14:trim/>",
            "<p:extLst/><p14:bmkLst/>",
        ] {
            assert!(
                parse_slide_media(&slide_xml(
                    SlideMediaConformance::Transitional,
                    &picture(children),
                ))
                .is_err(),
                "accepted invalid media children: {children}"
            );
        }
    }

    #[test]
    fn opaque_media_extension_constructor_is_bounded_and_typed() {
        assert!(MediaExtensionList::parse(b"<not-an-extension-list/>").is_err());
        assert!(MediaExtensionList::parse(&vec![b' '; MAX_MEDIA_EXTENSION_XML_BYTES + 1]).is_err());
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
    fn repeated_media_targets_share_one_immutable_payload_allocation() {
        let (mut package, uri) = package(SlideMediaConformance::Transitional);
        let mut expected = value();
        let mut second = expected.pictures[0].clone();
        second.shape_id = 5;
        second.name = "sample-copy.mp4".into();
        second.relationship_id = "rIdVideo2".into();
        second.poster.as_mut().unwrap().relationship_id = "rIdPoster2".into();
        second
            .office_extension
            .as_mut()
            .unwrap()
            .embed_relationship_id = Some("rIdMedia2".into());
        expected.pictures.push(second);

        store_slide_media(
            &mut package,
            &uri,
            &expected,
            SlideMediaConformance::Transitional,
        )
        .unwrap();
        let loaded = load_slide_media(&package, &uri).unwrap();
        let first = &loaded.pictures[0].resource.as_ref().unwrap().data;
        let second = &loaded.pictures[1].resource.as_ref().unwrap().data;
        assert!(first.shares_with(second));
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

    #[test]
    fn bookmark_uniqueness_compares_represented_time() {
        let mut value = value();
        let extension = value.pictures[0].office_extension.as_mut().unwrap();
        extension.bookmarks = vec![
            MediaBookmark {
                name: Some("first".into()),
                time: Some(Offset::parse("1s").unwrap()),
            },
            MediaBookmark {
                name: Some("second".into()),
                time: Some(Offset::parse("1000ms").unwrap()),
            },
        ];

        assert!(write_slide_media_pictures(&value, SlideMediaConformance::Transitional).is_err());
    }

    #[test]
    fn non_image_resource_requires_typed_media_kind() {
        let value = value();
        let resource = value.pictures[0].resource.as_ref().unwrap();

        assert!(matches!(
            resource_uri(resource, false, None),
            Err(Error::Invalid(message)) if message.contains("requires a media kind")
        ));
    }

    #[test]
    fn escaped_media_output_is_preflighted_against_a_typed_budget() {
        let mut value = value();
        value.pictures[0]
            .office_extension
            .as_mut()
            .unwrap()
            .bookmarks[0]
            .name = Some("\"".repeat(1_024));
        validate_value(&value, false).unwrap();

        let maximum = 2_048;
        let mut output = BoundedXml::with_limit(maximum);
        let error = write_picture(
            &mut output,
            &value.pictures[0],
            SlideMediaConformance::Transitional,
        )
        .unwrap_err();

        assert!(matches!(
            error,
            Error::Limit {
                resource: "slide media serialized XML bytes",
                limit,
            } if limit == maximum
        ));
        assert!(output.bytes.len() <= maximum);
    }

    #[test]
    fn media_output_budget_is_aggregate_across_pictures() {
        let picture = value().pictures.pop().unwrap();
        let mut single = BoundedXml::new();
        write_picture(&mut single, &picture, SlideMediaConformance::Transitional).unwrap();
        let single_len = single.bytes.len();
        let maximum = single_len + single_len / 2;

        let mut output = BoundedXml::with_limit(maximum);
        write_picture(&mut output, &picture, SlideMediaConformance::Transitional).unwrap();
        let error =
            write_picture(&mut output, &picture, SlideMediaConformance::Transitional).unwrap_err();

        assert!(matches!(
            error,
            Error::Limit {
                resource: "slide media serialized XML bytes",
                limit,
            } if limit == maximum
        ));
        assert!(output.bytes.len() >= single_len);
        assert!(output.bytes.len() <= maximum);
    }
}
