//! OLE equation objects: MathType (`Equation.DSMT4`) and Equation Editor
//! 3.0 (`Equation.3`) objects embedded in a slide.
//!
//! A slide shows an embedded OLE object through a `p:graphicFrame` whose
//! `a:graphicData` (URI `…/presentationml/2006/ole`) holds a `p:oleObj`
//! (ECMA-376 Part 1, §19.3.2, embedded objects) naming its ProgID, its
//! embedding part
//! (`r:id`, an `oleObject` relationship to `…/embeddings/oleObjectN.bin`)
//! and, in a `p:pic` child, the preview picture readers show. PowerPoint
//! 2010 and later write the object twice in an `mc:AlternateContent`: the
//! `mc:Choice Requires="v"` branch names a shape of the slide's legacy VML
//! drawing (`spid`), and the `mc:Fallback` branch holds the `p:pic` preview
//! (ECMA-376 Part 3 for markup compatibility). An object counts as an
//! equation when its ProgID is `Equation.3` or starts with
//! `Equation.DSMT`, as `litchi_equation::MtefTarget::from_prog_id`
//! decides.
//!
//! [`OleEquation`] lists one such object; with the `equation` feature its
//! `OleEquation::equation_native` reads the `Equation Native` stream of
//! the embedded compound file on first use. The edit verbs
//! `opened::Edit::{add_ole_equation, set_ole_equation, remove_ole_equation}`
//! of the same feature write them.

use std::ops::Range;
#[cfg(feature = "equation")]
use std::sync::{Arc, OnceLock};

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{MAX_DEPTH, MAX_EQUATIONS, MAX_NODES, MCE, PML, STRICT_PML, invalid};
use crate::{Error, Result};

const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const RELATIONSHIPS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
/// The `a:graphicData` URI of an OLE object frame (ECMA-376 Part 1,
/// §19.3.2), transitional and strict.
#[cfg_attr(
    not(feature = "equation"),
    allow(dead_code, reason = "the edits use it with the equation feature")
)]
pub(crate) const OLE_URI: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";
#[cfg_attr(
    not(feature = "equation"),
    allow(dead_code, reason = "the edits use it with the equation feature")
)]
pub(crate) const STRICT_OLE_URI: &str = "http://purl.oclc.org/ooxml/presentationml/ole";

/// Whether `prog_id` names an equation object: `Equation.3`, or a ProgID
/// starting with `Equation.DSMT`, ignoring case as COM does.
#[must_use]
pub fn is_equation_prog_id(prog_id: &str) -> bool {
    prog_id.eq_ignore_ascii_case("Equation.3")
        || prog_id
            .get(..13)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("Equation.DSMT"))
}

/// Which branch of an `mc:AlternateContent` a `p:oleObj` sits in.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Branch {
    /// Directly in the frame's `a:graphicData`.
    Direct,
    /// In an `mc:Choice` (PowerPoint's VML branch).
    Choice,
    /// In an `mc:Fallback` (the branch with the `p:pic` preview).
    Fallback,
}

/// One `p:oleObj` element of a frame.
#[cfg_attr(
    not(feature = "equation"),
    allow(dead_code, reason = "the edits use it with the equation feature")
)]
#[derive(Debug, Clone)]
pub(crate) struct ObjectSite {
    /// The whole element.
    pub(crate) element: Range<usize>,
    /// Its start tag.
    pub(crate) tag: Range<usize>,
    pub(crate) branch: Branch,
    pub(crate) prog_id: Option<String>,
    /// Its `r:id`.
    pub(crate) relationship: Option<String>,
    /// Its legacy VML shape (`spid`).
    pub(crate) spid: Option<String>,
    /// Whether it has a `p:embed` child (an embedded, not linked, object).
    pub(crate) embedded: bool,
    /// The start tag of each `a:blip` of its `p:pic` and the blip's
    /// `r:embed`.
    pub(crate) blips: Vec<(Range<usize>, Option<String>)>,
}

/// One OLE object frame of a slide.
#[cfg_attr(
    not(feature = "equation"),
    allow(dead_code, reason = "the edits use it with the equation feature")
)]
#[derive(Debug, Clone)]
pub(crate) struct FrameSite {
    /// The whole `p:graphicFrame` element.
    pub(crate) range: Range<usize>,
    /// Whether the frame itself sits in a branch of an
    /// `mc:AlternateContent` (not PowerPoint's form; edits are refused).
    pub(crate) in_block: bool,
    pub(crate) shape_id: Option<u32>,
    pub(crate) name: Option<String>,
    /// The frame's `p:xfrm` offset and extent.
    pub(crate) offset: Option<(i64, i64)>,
    pub(crate) extent: Option<(i64, i64)>,
    /// The `a:ext` start tags an edit resizes: the frame's and those of
    /// the preview pictures' `p:spPr`.
    pub(crate) extent_tags: Vec<Range<usize>>,
    /// The `mc:AlternateContent` inside the frame's `a:graphicData`.
    pub(crate) block: Option<Range<usize>>,
    pub(crate) objects: Vec<ObjectSite>,
}

impl FrameSite {
    /// The ProgID of the frame's object (the first `p:oleObj` naming one).
    pub(crate) fn prog_id(&self) -> Option<&str> {
        self.objects
            .iter()
            .find_map(|object| object.prog_id.as_deref())
    }

    /// The embedding relationship (the first `p:oleObj` naming one).
    pub(crate) fn relationship(&self) -> Option<&str> {
        self.objects
            .iter()
            .find_map(|object| object.relationship.as_deref())
    }

    /// The preview picture relationship: the first `a:blip` naming one.
    pub(crate) fn preview(&self) -> Option<&str> {
        self.objects
            .iter()
            .flat_map(|object| object.blips.iter())
            .find_map(|(_, embed)| embed.as_deref())
    }

    /// The legacy VML shape IDs the frame names.
    pub(crate) fn spids(&self) -> Vec<&str> {
        let mut out: Vec<&str> = Vec::new();
        for spid in self
            .objects
            .iter()
            .filter_map(|object| object.spid.as_deref())
        {
            if !out.contains(&spid) {
                out.push(spid);
            }
        }
        out
    }

    /// Whether the frame holds an equation object.
    fn is_equation(&self) -> bool {
        self.prog_id().is_some_and(is_equation_prog_id)
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Space {
    Pml,
    Dml,
    Mce,
    Other,
}

fn space(namespace: &ResolveResult<'_>) -> Space {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == PML || *value == STRICT_PML => {
            Space::Pml
        },
        ResolveResult::Bound(Namespace(value)) if *value == DML || *value == STRICT_DML => {
            Space::Dml
        },
        ResolveResult::Bound(Namespace(value)) if *value == MCE => Space::Mce,
        _ => Space::Other,
    }
}

/// What an open element of a frame is, for the scan.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Tag {
    Frame,
    FrameTransform,
    Block,
    Choice,
    Fallback,
    Object,
    PictureProperties,
    PictureTransform,
    Other,
}

/// An attribute's unescaped value.
fn value(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    element
        .checked_attributes()
        .flatten()
        .find(|attribute| attribute.key.as_ref() == name)
        .and_then(|attribute| {
            attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
                .ok()
                .map(std::borrow::Cow::into_owned)
        })
}

/// The value of a relationship-namespace attribute `local` (`r:id`,
/// `r:embed`).
fn relationship_value(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Option<String> {
    for attribute in element.checked_attributes().flatten() {
        if attribute.key.local_name().as_ref() != local {
            continue;
        }
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == RELATIONSHIPS || uri == STRICT_RELATIONSHIPS)
        {
            return attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
                .ok()
                .map(std::borrow::Cow::into_owned);
        }
    }
    None
}

fn integer(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, name: &[u8]) -> Option<i64> {
    value(reader, element, name)?.trim().parse().ok()
}

/// Every OLE equation frame of a slide's XML, in document order. A frame
/// that sits in the Fallback of a frame-level `mc:AlternateContent` whose
/// Choice already holds a frame is the same object and is not listed again.
///
/// # Errors
///
/// Returns an error for malformed XML or exceeded bounds.
pub(crate) fn scan(xml: &[u8]) -> Result<Vec<FrameSite>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let position = |reader: &NsReader<&[u8]>| {
        usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("the slide is too large"))
    };
    // Every open element: what it is, and where it starts. Outside a frame
    // only the markup-compatibility elements are told apart.
    let mut stack: Vec<(Tag, usize)> = Vec::new();
    let mut frame: Option<FrameSite> = None;
    // For each open frame-level block: whether a branch held a frame.
    let mut blocks: Vec<(usize, bool)> = Vec::new();
    let mut out = Vec::new();
    let mut nodes = 0usize;
    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let space = space(&namespace);
        let end = position(&reader)?;
        nodes += 1;
        if nodes > MAX_NODES {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                observed: nodes,
                maximum: MAX_NODES,
            });
        }
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                let local = element.local_name();
                let local = local.as_ref();
                let parent = stack.last().map_or(Tag::Other, |(tag, _)| *tag);
                let tag = if let Some(site) = frame.as_mut() {
                    let in_object = stack.iter().any(|(tag, _)| *tag == Tag::Object);
                    match (space, local, parent) {
                        (Space::Pml, b"cNvPr", _) if site.shape_id.is_none() && !in_object => {
                            site.shape_id = integer(&reader, element, b"id")
                                .and_then(|id| u32::try_from(id).ok());
                            site.name = value(&reader, element, b"name");
                            Tag::Other
                        },
                        (Space::Pml, b"xfrm", Tag::Frame) => Tag::FrameTransform,
                        (Space::Dml, b"off", Tag::FrameTransform) => {
                            site.offset = integer(&reader, element, b"x")
                                .zip(integer(&reader, element, b"y"));
                            Tag::Other
                        },
                        (Space::Dml, b"ext", Tag::FrameTransform) => {
                            site.extent = integer(&reader, element, b"cx")
                                .zip(integer(&reader, element, b"cy"));
                            site.extent_tags.push(start..end);
                            Tag::Other
                        },
                        (Space::Mce, b"AlternateContent", _) if !in_object => {
                            if site.block.is_none() {
                                site.block = Some(start..end);
                            }
                            Tag::Block
                        },
                        (Space::Mce, b"Choice", Tag::Block) => Tag::Choice,
                        (Space::Mce, b"Fallback", Tag::Block) => Tag::Fallback,
                        (Space::Pml, b"oleObj", _) if !in_object => {
                            let branch = match parent {
                                Tag::Choice => Branch::Choice,
                                Tag::Fallback => Branch::Fallback,
                                _ => Branch::Direct,
                            };
                            site.objects.push(ObjectSite {
                                element: start..end,
                                tag: start..end,
                                branch,
                                prog_id: value(&reader, element, b"progId"),
                                relationship: relationship_value(&reader, element, b"id"),
                                spid: value(&reader, element, b"spid"),
                                embedded: false,
                                blips: Vec::new(),
                            });
                            Tag::Object
                        },
                        (Space::Pml, b"embed", Tag::Object) => {
                            if let Some(object) = site.objects.last_mut() {
                                object.embedded = true;
                            }
                            Tag::Other
                        },
                        (Space::Pml, b"spPr", _) if in_object => Tag::PictureProperties,
                        (Space::Dml, b"xfrm", Tag::PictureProperties) => Tag::PictureTransform,
                        (Space::Dml, b"ext", Tag::PictureTransform) => {
                            site.extent_tags.push(start..end);
                            Tag::Other
                        },
                        (Space::Dml, b"blip", _) if in_object => {
                            let embed = relationship_value(&reader, element, b"embed");
                            if let Some(object) = site.objects.last_mut() {
                                object.blips.push((start..end, embed));
                            }
                            Tag::Other
                        },
                        _ => Tag::Other,
                    }
                } else if space == Space::Pml && local == b"graphicFrame" {
                    let in_block = matches!(parent, Tag::Choice | Tag::Fallback);
                    // A frame in the Fallback of a block whose Choice held
                    // a frame repeats that frame.
                    let repeated =
                        parent == Tag::Fallback && blocks.last().is_some_and(|(_, held)| *held);
                    if in_block && let Some((_, held)) = blocks.last_mut() {
                        *held = true;
                    }
                    if !repeated && !empty {
                        frame = Some(FrameSite {
                            range: start..end,
                            in_block,
                            shape_id: None,
                            name: None,
                            offset: None,
                            extent: None,
                            extent_tags: Vec::new(),
                            block: None,
                            objects: Vec::new(),
                        });
                    }
                    Tag::Frame
                } else {
                    match (space, local, parent) {
                        (Space::Mce, b"AlternateContent", _) => {
                            if !empty {
                                blocks.push((stack.len(), false));
                            }
                            Tag::Block
                        },
                        (Space::Mce, b"Choice", Tag::Block) => Tag::Choice,
                        (Space::Mce, b"Fallback", Tag::Block) => Tag::Fallback,
                        _ => Tag::Other,
                    }
                };
                if empty {
                    if tag == Tag::Object
                        && let Some(site) = frame.as_mut()
                        && let Some(object) = site.objects.last_mut()
                    {
                        object.element = start..end;
                    }
                } else {
                    stack.push((tag, start));
                    if stack.len() > MAX_DEPTH {
                        return Err(invalid("the slide is nested too deeply"));
                    }
                }
            },
            Event::End(_) => {
                let (tag, opened) = stack.pop().ok_or_else(|| invalid("unbalanced slide XML"))?;
                match tag {
                    Tag::Object => {
                        if let Some(site) = frame.as_mut()
                            && let Some(object) = site.objects.last_mut()
                        {
                            object.element = opened..end;
                        }
                    },
                    Tag::Block => {
                        if let Some(site) = frame.as_mut() {
                            if let Some(block) = site.block.as_mut()
                                && block.start == opened
                            {
                                block.end = end;
                            }
                        } else if blocks
                            .last()
                            .is_some_and(|(depth, _)| *depth == stack.len())
                        {
                            blocks.pop();
                        }
                    },
                    Tag::Frame => {
                        if let Some(mut site) = frame.take()
                            && site.range.start == opened
                        {
                            site.range.end = end;
                            if site.is_equation() {
                                if out.len() >= MAX_EQUATIONS {
                                    return Err(Error::Limit {
                                        kind: crate::opened::LimitKind::Codec(
                                            crate::opened::CodecLimitKind::Equations,
                                        ),
                                        observed: out.len().saturating_add(1),
                                        maximum: MAX_EQUATIONS,
                                    });
                                }
                                out.push(site);
                            }
                        }
                    },
                    _ => {},
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(out)
}

/// Why an OLE equation's `Equation Native` stream cannot be read. The
/// equation is still listed; only its content is missing.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Unreadable {
    reason: String,
}

impl Unreadable {
    #[cfg_attr(
        not(feature = "equation"),
        allow(dead_code, reason = "the content is read with the equation feature")
    )]
    pub(crate) fn new(reason: impl Into<String>) -> Self {
        Self {
            reason: reason.into(),
        }
    }

    /// What is wrong with the embedding.
    #[must_use]
    pub fn reason(&self) -> &str {
        &self.reason
    }
}

impl std::fmt::Display for Unreadable {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "the equation cannot be read: {}", self.reason)
    }
}

impl std::error::Error for Unreadable {}

/// Where an equation's embedded compound file comes from.
#[cfg(feature = "equation")]
#[derive(Clone)]
pub(crate) enum Embedding {
    /// The embedding cannot be reached, and why.
    Unavailable(String),
    /// The part's bytes, shared with the package.
    Bytes(Arc<Vec<u8>>),
    /// Read from the source package on first use.
    Deferred(Arc<dyn Fn() -> Result<Vec<u8>> + Send + Sync>),
}

/// An equation's `Equation Native` stream, read on first use.
#[derive(Clone)]
pub(crate) struct Content {
    #[cfg(feature = "equation")]
    embedding: Embedding,
    #[cfg(feature = "equation")]
    native: Arc<OnceLock<std::result::Result<Arc<[u8]>, Unreadable>>>,
}

impl Content {
    #[cfg(feature = "equation")]
    pub(crate) fn new(embedding: Embedding) -> Self {
        Self {
            embedding,
            native: Arc::new(OnceLock::new()),
        }
    }

    #[cfg(not(feature = "equation"))]
    pub(crate) const fn new() -> Self {
        Self {}
    }

    #[cfg(feature = "equation")]
    fn native(&self) -> std::result::Result<Arc<[u8]>, Unreadable> {
        self.native
            .get_or_init(|| {
                let bytes = match &self.embedding {
                    Embedding::Unavailable(reason) => return Err(Unreadable::new(reason.clone())),
                    Embedding::Bytes(bytes) => litchi_equation::equation_object_native(bytes),
                    Embedding::Deferred(load) => match load() {
                        Ok(bytes) => litchi_equation::equation_object_native(&bytes),
                        Err(error) => return Err(Unreadable::new(error.to_string())),
                    },
                };
                bytes
                    .map(Arc::from)
                    .map_err(|error| Unreadable::new(error.to_string()))
            })
            .clone()
    }
}

/// One OLE equation object of a slide.
///
/// Listing reads no compound file: `OleEquation::equation_native` reads
/// the embedding's `Equation Native` stream on first use, once, and an
/// object whose embedding cannot be read is still listed with the reason.
#[derive(Clone)]
pub struct OleEquation {
    shape_id: Option<u32>,
    name: Option<String>,
    prog_id: String,
    offset: Option<(i64, i64)>,
    extent: Option<(i64, i64)>,
    embedding: Option<String>,
    preview: Option<String>,
    legacy_drawing: bool,
    #[cfg_attr(
        not(feature = "equation"),
        allow(dead_code, reason = "the content is read with the equation feature")
    )]
    content: Content,
}

impl std::fmt::Debug for OleEquation {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("OleEquation")
            .field("shape_id", &self.shape_id)
            .field("prog_id", &self.prog_id)
            .field("embedding", &self.embedding)
            .finish_non_exhaustive()
    }
}

impl OleEquation {
    /// The `cNvPr` ID of the object's frame.
    #[must_use]
    pub const fn shape_id(&self) -> Option<u32> {
        self.shape_id
    }

    /// The frame's name (`cNvPr@name`).
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// The object's ProgID: `Equation.DSMT4` (or another `Equation.DSMT*`)
    /// for MathType, `Equation.3` for Equation Editor 3.0.
    #[must_use]
    pub fn prog_id(&self) -> &str {
        &self.prog_id
    }

    /// The frame's offset (`x`, `y` in EMU), when it has one.
    #[must_use]
    pub const fn offset(&self) -> Option<(i64, i64)> {
        self.offset
    }

    /// The frame's extent (`cx`, `cy` in EMU), when it has one.
    #[must_use]
    pub const fn extent(&self) -> Option<(i64, i64)> {
        self.extent
    }

    /// The name of the embedded compound-file part, when the object's
    /// relationship names one.
    #[must_use]
    pub fn embedding_part(&self) -> Option<&str> {
        self.embedding.as_deref()
    }

    /// The name of the preview picture part, when the object has one.
    #[must_use]
    pub fn preview_part(&self) -> Option<&str> {
        self.preview.as_deref()
    }

    /// Whether the object also names a shape of the slide's legacy VML
    /// drawing (PowerPoint's `mc:Choice Requires="v"` form).
    #[must_use]
    pub const fn has_legacy_drawing(&self) -> bool {
        self.legacy_drawing
    }

    /// The `Equation Native` stream, exactly as stored: the 28-byte
    /// `EQNOLEFILEHDR` followed by MTEF.
    ///
    /// Requires the `equation` feature.
    ///
    /// # Errors
    ///
    /// Returns [`Unreadable`] when the embedding is missing, is not a
    /// compound file, holds no such stream, or exceeds
    /// [`litchi_equation::MAX_EQUATION_OBJECT_BYTES`] or
    /// [`litchi_equation::MAX_EQUATION_NATIVE_BYTES`].
    #[cfg(feature = "equation")]
    pub fn equation_native(&self) -> std::result::Result<Arc<[u8]>, Unreadable> {
        self.content.native()
    }

    /// The equation as LaTeX.
    ///
    /// Requires the `equation` feature.
    ///
    /// # Errors
    ///
    /// Returns the conversion error when the MTEF does not read within
    /// `limits`, and [`litchi_equation::Error::Malformed`] when the stream
    /// is [`Unreadable`].
    #[cfg(feature = "equation")]
    pub fn latex(&self, limits: &litchi_equation::Limits) -> litchi_equation::Result<String> {
        let native = self
            .equation_native()
            .map_err(|error| litchi_equation::Error::Malformed {
                format: litchi_equation::Format::Mtef,
                message: error.reason().to_owned(),
            })?;
        litchi_equation::mtef_to_latex(&native, limits)
    }
}

/// What a slide relationship resolves to, for the listing.
pub(crate) enum Resolved {
    /// An internal part, by name.
    Part(String),
    /// An external target, a missing relationship or one of another type.
    Other(&'static str),
}

/// The OLE equations of a slide's XML, in document order: `resolve` maps a
/// relationship ID of the slide to its target, `content` builds the lazy
/// content of an embedding part.
pub(crate) fn list(
    xml: &[u8],
    resolve: impl Fn(&str) -> Result<Resolved>,
    content: impl Fn(std::result::Result<&str, &'static str>) -> Content,
) -> Result<Vec<OleEquation>> {
    let mut out = Vec::new();
    for site in scan(xml)? {
        let embedding = match site.relationship() {
            Some(id) => match resolve(id)? {
                Resolved::Part(name) => Ok(name),
                Resolved::Other(reason) => Err(reason),
            },
            None => Err("the object names no embedding"),
        };
        let preview = match site.preview() {
            Some(id) => match resolve(id)? {
                Resolved::Part(name) => Some(name),
                Resolved::Other(_) => None,
            },
            None => None,
        };
        let linked = !site.objects.iter().any(|object| object.embedded);
        let embedding = if linked {
            Err("the object is linked, not embedded")
        } else {
            embedding
        };
        out.push(OleEquation {
            shape_id: site.shape_id,
            name: site.name.clone(),
            prog_id: site.prog_id().unwrap_or_default().to_owned(),
            offset: site.offset,
            extent: site.extent,
            content: content(embedding.as_deref().map_err(|reason| *reason)),
            embedding: embedding.ok(),
            preview,
            legacy_drawing: !site.spids().is_empty(),
        });
    }
    Ok(out)
}

/// Resolve a slide relationship for [`list`]: an internal target, by name.
pub(crate) fn resolve(relationship: Option<&litchi_opc::Relationship>) -> Result<Resolved> {
    let Some(relationship) = relationship else {
        return Ok(Resolved::Other("the object's relationship is missing"));
    };
    if relationship.is_external() {
        return Ok(Resolved::Other("the object's relationship is external"));
    }
    Ok(Resolved::Part(relationship.target_partname()?.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIXTURE: &[u8] = include_bytes!(
        "../../../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf169088.pptx"
    );

    #[test]
    fn prog_ids_name_equations() {
        for yes in [
            "Equation.DSMT4",
            "equation.dsmt6",
            "Equation.3",
            "EQUATION.3",
        ] {
            assert!(is_equation_prog_id(yes), "{yes}");
        }
        for no in [
            "Equation.2",
            "Equation.DSM",
            "Excel.Sheet.12",
            "",
            "Equation",
        ] {
            assert!(!is_equation_prog_id(no), "{no}");
        }
    }

    #[test]
    fn scans_powerpoints_markup_compatibility_form() {
        let package = crate::Package::from_bytes(FIXTURE).unwrap();
        let presentation = package.presentation().unwrap();
        let slide = presentation.slide(0).unwrap().unwrap();
        let xml = slide.part().part().blob().to_vec();
        let sites = scan(&xml).unwrap();
        assert_eq!(sites.len(), 1);
        let site = &sites[0];
        assert_eq!(site.shape_id, Some(295_958));
        assert_eq!(site.name.as_deref(), Some("Object 22"));
        assert_eq!(site.offset, Some((3_806_825, 5_440_363)));
        assert_eq!(site.extent, Some((1_101_725, 752_475)));
        assert!(!site.in_block);
        assert_eq!(site.objects.len(), 2);
        assert_eq!(site.objects[0].branch, Branch::Choice);
        assert_eq!(site.objects[1].branch, Branch::Fallback);
        assert_eq!(site.prog_id(), Some("Equation.DSMT4"));
        assert_eq!(site.relationship(), Some("rId3"));
        assert_eq!(site.preview(), Some("rId4"));
        assert_eq!(site.spids(), ["_x0000_s295990"]);
        // The frame's and the preview picture's extents.
        assert_eq!(site.extent_tags.len(), 2);
        let block = site.block.clone().unwrap();
        assert!(xml[block].starts_with(b"<mc:AlternateContent"));
        assert!(xml[site.range.clone()].ends_with(b"</p:graphicFrame>"));
        for object in &site.objects {
            assert!(xml[object.element.clone()].ends_with(b"</p:oleObj>"));
        }
    }

    #[test]
    fn other_objects_and_repeated_frames_are_not_listed() {
        let frame = |id: u32, prog: &str| {
            format!(
                r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{id}" name="o"/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="{OLE_URI}"><p:oleObj progId="{prog}" r:id="rId1"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>"#
            )
        };
        let xml = format!(
            r#"<p:sld xmlns:p="{}" xmlns:a="{}" xmlns:r="{}" xmlns:mc="{}"><p:cSld><p:spTree>{}<mc:AlternateContent><mc:Choice Requires="v">{}</mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#,
            std::str::from_utf8(PML).unwrap(),
            std::str::from_utf8(DML).unwrap(),
            std::str::from_utf8(RELATIONSHIPS).unwrap(),
            std::str::from_utf8(MCE).unwrap(),
            frame(2, "Excel.Sheet.12"),
            frame(3, "Equation.3"),
            frame(3, "Equation.3"),
        );
        let sites = scan(xml.as_bytes()).unwrap();
        assert_eq!(sites.len(), 1);
        assert_eq!(sites[0].shape_id, Some(3));
        assert!(sites[0].in_block);
        assert!(scan(b"<p:sld><unclosed></p:sld>").is_err());
    }
}
