//! Byte-preserving flat `OpenDocument` Image snapshots and bounded frame edits.
#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The file keeps public snapshot/transaction types before their private XML machinery."
)]

use crate::{
    frame::{Frame, Properties},
    map::{Area, AreaKind, AreaProperties, ImageMap},
    source::Source,
};
use litchi_core::{Error, FileFormat, Result};
use litchi_odf_common::{compact_xml, media};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{ops::Range, sync::Arc};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_FRAMES: usize = 1_000_000;
const MAX_MAP_AREAS: usize = 100_000;
const MAX_MAP_TEXT_BYTES: usize = 64 * 1024;

#[derive(Debug)]
struct State {
    bytes: Vec<u8>,
    root: Root,
    frames: Vec<Frame>,
    sites: Vec<FrameSite>,
}

#[derive(Clone, Debug)]
struct FrameSite {
    frame_tag: Option<Range<usize>>,
    attributes: FrameAttributeSites,
    properties: Properties,
    image_tag: Range<usize>,
    href_attribute: Option<String>,
    binary_contents: Option<Range<usize>>,
}

#[derive(Clone, Debug, Default)]
struct FrameAttributeSites {
    name: Option<String>,
    style_name: Option<String>,
    text_style_name: Option<String>,
    layer: Option<String>,
    z_index: Option<String>,
    transform: Option<String>,
    anchor_type: Option<String>,
    x: Option<String>,
    y: Option<String>,
    width: Option<String>,
    height: Option<String>,
    relative_width: Option<String>,
    relative_height: Option<String>,
}

/// An immutable, byte-preserving flat ODI snapshot.
#[derive(Clone, Debug)]
pub struct FlatImage(Arc<State>);

impl FlatImage {
    /// Opens a flat ODI document and inventories its inert image frames.
    ///
    /// # Errors
    ///
    /// Returns an error for the wrong family, malformed XML, unsupported image
    /// sources, or exceeded input limits.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        if bytes.len() > MAX_BYTES {
            return Err(invalid("flat ODI exceeds the input size limit"));
        }
        if litchi_odf_common::detect::flat(&bytes) != Some(FileFormat::Odi) {
            return Err(invalid("input is not a flat ODI document"));
        }
        Self::from_root_bytes(bytes, Root::Flat)
    }

    /// Parses a packaged ODI `content.xml` part with the same bounded frame
    /// inventory and lossless edit-site scanner used for flat documents.
    pub(crate) fn from_content_xml(bytes: Vec<u8>) -> Result<Self> {
        if bytes.len() > MAX_BYTES {
            return Err(invalid("ODI content.xml exceeds the input size limit"));
        }
        Self::from_root_bytes(bytes, Root::Content)
    }

    fn from_root_bytes(bytes: Vec<u8>, root: Root) -> Result<Self> {
        Ok(Self(Arc::new(parse(bytes, root)?)))
    }

    /// Returns the inert frames in source order.
    #[must_use]
    pub fn frames(&self) -> &[Frame] {
        &self.0.frames
    }

    /// Returns the document's single normative image frame.
    #[must_use]
    pub fn frame(&self) -> Option<&Frame> {
        self.0.frames.first()
    }

    /// Returns the original flat XML bytes exactly.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.0.bytes
    }

    /// Starts a source-bound transaction without changing this snapshot.
    #[must_use]
    pub fn transaction(&self) -> FlatImageTransaction {
        FlatImageTransaction {
            source: self.clone(),
            changes: Vec::new(),
        }
    }

    /// Consumes the snapshot and returns its exact flat XML bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        Arc::try_unwrap(self.0).map_or_else(|state| state.bytes.clone(), |state| state.bytes)
    }
}

pub(crate) fn validate_content_xml(xml: &str) -> Result<()> {
    if xml.len() > MAX_BYTES {
        return Err(invalid("ODI content.xml exceeds the input size limit"));
    }
    validate_structure(xml, Root::Content)?;
    let images = media::scan_content(xml)?;
    if images.len() != 1 {
        return Err(invalid("ODI content.xml must contain exactly one image"));
    }
    match &images[0].source {
        media::Source::Inline { .. }
        | media::Source::Linked { .. }
        | media::Source::PackagePart { .. }
        | media::Source::MissingPackagePart { .. } => Ok(()),
        media::Source::Missing => Err(invalid("ODI draw:image has no image source")),
        _ => Err(invalid("ODI draw:image source is not supported")),
    }
}

/// A source-bound mutable draft of flat ODI frame metadata.
pub struct FlatImageTransaction {
    source: FlatImage,
    changes: Vec<FrameChange>,
}

impl FlatImageTransaction {
    /// Stages a replacement for the document frame's optional name.
    ///
    /// # Errors
    ///
    /// Returns an error if the frame has no losslessly editable owner.
    pub fn set_name(&mut self, name: Option<String>) -> Result<()> {
        self.set_frame_name(0, name)
    }

    /// Stages a replacement for the document's linked or inline image source.
    ///
    /// # Errors
    ///
    /// Returns an error if changing source representation would be lossy.
    pub fn set_image_source(&mut self, source: Source) -> Result<()> {
        self.set_source(0, source)
    }

    /// Stages a replacement for a frame's optional `draw:name` attribute.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or uneditable owner.
    pub fn set_frame_name(&mut self, frame: usize, name: Option<String>) -> Result<()> {
        if self.site(frame)?.frame_tag.is_none() {
            return Err(invalid(
                "flat ODI image frame has no editable draw:frame owner",
            ));
        }
        self.change_mut(frame)?.after.set_name(name);
        self.remove_noops();
        Ok(())
    }

    /// Stages the optional graphic style reference.
    pub fn set_style_name(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.change_mut(frame)?.after.set_style_name(value);
        self.remove_noops();
        Ok(())
    }

    /// Stages the optional paragraph style reference used by frame text.
    pub fn set_text_style_name(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.change_mut(frame)?.after.set_text_style_name(value);
        self.remove_noops();
        Ok(())
    }

    /// Stages the optional drawing layer name.
    pub fn set_layer(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.change_mut(frame)?.after.set_layer(value);
        self.remove_noops();
        Ok(())
    }

    /// Stages the optional non-negative stacking order.
    pub fn set_z_index(&mut self, frame: usize, value: Option<u32>) -> Result<()> {
        self.change_mut(frame)?.after.set_z_index(value);
        self.remove_noops();
        Ok(())
    }

    /// Stages the optional lexical drawing transform.
    pub fn set_transform(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.change_mut(frame)?.after.set_transform(value);
        self.remove_noops();
        Ok(())
    }

    /// Stages the optional text anchoring mode.
    pub fn set_anchor_type(&mut self, frame: usize, value: Option<String>) -> Result<()> {
        self.change_mut(frame)?.after.set_anchor_type(value);
        self.remove_noops();
        Ok(())
    }

    /// Stages lexical position and size values without unit normalization.
    pub fn set_geometry(
        &mut self,
        frame: usize,
        x: Option<String>,
        y: Option<String>,
        width: Option<String>,
        height: Option<String>,
    ) -> Result<()> {
        self.change_mut(frame)?
            .after
            .set_geometry(x, y, width, height);
        self.remove_noops();
        Ok(())
    }

    /// Stages lexical relative width and height values.
    pub fn set_relative_size(
        &mut self,
        frame: usize,
        width: Option<String>,
        height: Option<String>,
    ) -> Result<()> {
        self.change_mut(frame)?
            .after
            .set_relative_size(width, height);
        self.remove_noops();
        Ok(())
    }

    /// Stages replacement of an existing linked URI or inline binary payload.
    ///
    /// Cross-kind source changes refuse rather than reconstructing unknown XML.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or lossy representation change.
    pub fn set_source(&mut self, frame: usize, source: Source) -> Result<()> {
        let before_source = self.frame(frame)?.source().clone();
        if before_source == source {
            self.change_mut(frame)?.after.set_source(before_source);
            self.remove_noops();
            return Ok(());
        }
        let site = self.site(frame)?;
        match (&before_source, &source) {
            (Source::Linked(_), Source::Linked(_)) if site.href_attribute.is_some() => {},
            (Source::Embedded(_), Source::Embedded(_)) if site.binary_contents.is_some() => {},
            (Source::Linked(_), Source::Linked(_)) => {
                return Err(invalid("flat ODI linked image has no editable xlink:href"));
            },
            (Source::Embedded(_), Source::Embedded(_)) => {
                return Err(invalid(
                    "flat ODI inline image has no editable binary-data content",
                ));
            },
            _ => {
                return Err(invalid(
                    "flat ODI source representation changes are not lossless",
                ));
            },
        }
        self.change_mut(frame)?.after.set_source(source);
        self.remove_noops();
        Ok(())
    }

    /// Atomically validates and publishes an immutable edited snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if staged XML cannot be rewritten and read back safely.
    pub fn commit(mut self) -> Result<FlatImageCommit> {
        if self.changes.is_empty() {
            return Ok(FlatImageCommit {
                snapshot: self.source.clone(),
                patch: FlatImagePatch {
                    source: self.source.clone(),
                    target: self.source,
                    changes: Vec::new(),
                },
            });
        }
        self.changes.sort_unstable_by_key(|change| change.frame);
        let mut edits = Vec::with_capacity(self.changes.len() * 2);
        for change in &self.changes {
            let site = self.site(change.frame)?;
            if frame_attributes_differ(&change.before, &change.after) {
                let span = site
                    .frame_tag
                    .as_ref()
                    .ok_or_else(|| invalid("flat ODI frame attribute site disappeared"))?;
                let raw = std::str::from_utf8(
                    self.source
                        .as_bytes()
                        .get(span.clone())
                        .ok_or_else(|| invalid("ODI frame tag span is invalid"))?,
                )
                .map_err(|error| invalid(format!("ODI frame tag is not UTF-8: {error}")))?;
                edits.push((
                    span.clone(),
                    rewrite_frame_tag(raw, &site.attributes, &change.before, &change.after)?,
                ));
            }
            if change.before.source() != change.after.source() {
                match change.after.source() {
                    Source::Linked(href) => edits.push((
                        site.image_tag.clone(),
                        rewrite_attribute(
                            self.source.as_bytes(),
                            &site.image_tag,
                            site.href_attribute.as_deref(),
                            "xlink:href",
                            Some(href),
                        )?,
                    )),
                    Source::Embedded(bytes) => edits.push((
                        site.binary_contents
                            .clone()
                            .ok_or_else(|| invalid("flat ODI inline image site disappeared"))?,
                        base64(bytes),
                    )),
                }
            }
        }
        let bytes = apply_edits(self.source.as_bytes(), edits)?;
        compact_xml::validate(&bytes)?;
        let snapshot = FlatImage::from_root_bytes(bytes, self.source.0.root)?;
        for change in &self.changes {
            let actual = snapshot
                .frames()
                .get(change.frame)
                .ok_or_else(|| invalid("flat ODI edit lost its selected frame"))?;
            if actual != &change.after {
                return Err(invalid("flat ODI edit failed semantic readback"));
            }
        }
        Ok(FlatImageCommit {
            snapshot: snapshot.clone(),
            patch: FlatImagePatch {
                source: self.source,
                target: snapshot,
                changes: self.changes,
            },
        })
    }

    fn frame(&self, index: usize) -> Result<&Frame> {
        self.source
            .frames()
            .get(index)
            .ok_or_else(|| invalid("flat ODI frame selector is out of bounds"))
    }

    fn site(&self, index: usize) -> Result<&FrameSite> {
        self.source
            .0
            .sites
            .get(index)
            .ok_or_else(|| invalid("flat ODI image site selector is out of bounds"))
    }

    fn change_mut(&mut self, index: usize) -> Result<&mut FrameChange> {
        if let Some(position) = self.changes.iter().position(|change| change.frame == index) {
            return Ok(&mut self.changes[position]);
        }
        let before = self.frame(index)?.clone();
        self.changes.push(FrameChange {
            frame: index,
            before: before.clone(),
            after: before,
        });
        self.changes
            .last_mut()
            .ok_or_else(|| invalid("flat ODI failed to stage a frame change"))
    }

    fn remove_noops(&mut self) {
        self.changes.retain(|change| change.before != change.after);
    }
}

/// A validated publication result and its reversible patch.
pub struct FlatImageCommit {
    snapshot: FlatImage,
    patch: FlatImagePatch,
}

impl FlatImageCommit {
    /// Returns the published immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &FlatImage {
        &self.snapshot
    }

    /// Returns the exact-source-checked patch.
    #[must_use]
    pub fn patch(&self) -> &FlatImagePatch {
        &self.patch
    }

    /// Consumes this result and returns the published snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> FlatImage {
        self.snapshot
    }
}

/// A source-checked reversible flat ODI metadata patch.
#[derive(Clone, Debug)]
pub struct FlatImagePatch {
    source: FlatImage,
    target: FlatImage,
    changes: Vec<FrameChange>,
}

impl FlatImagePatch {
    /// Returns whether the patch applies to this exact source byte sequence.
    #[must_use]
    pub fn is_applicable_to(&self, source: &FlatImage) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies the patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied source is not byte-identical.
    pub fn apply(&self, source: &FlatImage) -> Result<FlatImage> {
        if !self.is_applicable_to(source) {
            return Err(invalid("stale flat ODI patch source"));
        }
        Ok(self.target.clone())
    }

    /// Returns selector-bound metadata changes in source order.
    #[must_use]
    pub fn changes(&self) -> &[FrameChange] {
        &self.changes
    }

    /// Returns the patch that restores the exact source snapshot.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self.changes.iter().map(FrameChange::inverse).collect(),
        }
    }
}

/// One reversible frame metadata change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FrameChange {
    frame: usize,
    before: Frame,
    after: Frame,
}

impl FrameChange {
    #[must_use]
    pub fn frame(&self) -> usize {
        self.frame
    }
    #[must_use]
    pub fn before_name(&self) -> Option<&str> {
        self.before.name()
    }
    #[must_use]
    pub fn after_name(&self) -> Option<&str> {
        self.after.name()
    }
    #[must_use]
    pub fn before_source(&self) -> &Source {
        self.before.source()
    }
    #[must_use]
    pub fn after_source(&self) -> &Source {
        self.after.source()
    }

    /// Returns the complete source frame semantics.
    #[must_use]
    pub const fn before(&self) -> &Frame {
        &self.before
    }

    /// Returns the complete target frame semantics.
    #[must_use]
    pub const fn after(&self) -> &Frame {
        &self.after
    }

    fn inverse(&self) -> Self {
        Self {
            frame: self.frame,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum Root {
    Flat,
    Content,
}

fn parse(input_bytes: Vec<u8>, root: Root) -> Result<State> {
    let xml = std::str::from_utf8(&input_bytes)
        .map_err(|error| invalid(format!("ODI XML is not UTF-8: {error}")))?;
    validate_structure(xml, root)?;
    let images = media::scan_flat(xml)?;
    if images.len() > MAX_FRAMES {
        return Err(invalid("flat ODI frame count exceeds the limit"));
    }
    let sites = scan_sites(xml)?;
    if sites.len() != images.len() {
        return Err(invalid("flat ODI image sites cannot be matched losslessly"));
    }
    let mut frames = Vec::with_capacity(images.len());
    let image_map = scan_image_map(xml)?;
    for (image, site) in images.into_iter().zip(&sites) {
        let source = match &image.source {
            media::Source::Inline { bytes: payload, .. } => Source::Embedded(payload.clone()),
            media::Source::Linked { href }
            | media::Source::PackagePart { href, .. }
            | media::Source::MissingPackagePart { href, .. } => Source::Linked(href.clone()),
            media::Source::Missing => {
                return Err(invalid(
                    "flat ODI draw:image has no losslessly modeled source",
                ));
            },
            _ => return Err(invalid("flat ODI image source is not supported")),
        };
        let mut frame = Frame::from_scanned(source, &image);
        let mut properties = site.properties.clone();
        properties.image_map.clone_from(&image_map);
        frame.apply_properties(properties);
        frames.push(frame);
    }
    Ok(State {
        bytes: input_bytes,
        root,
        frames,
        sites,
    })
}

fn validate_structure(xml: &str, root: Root) -> Result<()> {
    // ODF 1.4 Part 3, sections 2.2.8 and 3.9: an image body owns one
    // `draw:frame`, and that frame owns one `draw:image`.
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut structure = Structure::default();
    loop {
        let (resolved_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODI XML: {error}")))?;
        let namespace = classify(&resolved_namespace);
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                structure.observe(namespace, &element, depth, false, root)?;
            },
            Event::Empty(element) => {
                structure.observe(namespace, &element, checked_depth(depth)?, true, root)?;
            },
            Event::End(_) => {
                structure.close(depth);
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODI XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODI XML")),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0
        || !structure.root_seen
        || !structure.body_seen
        || !structure.image_seen
        || !structure.frame_seen
        || structure.image_count != 1
    {
        return Err(invalid(
            "ODI requires office:body/office:image with one draw:frame containing one draw:image",
        ));
    }
    Ok(())
}

#[derive(Default)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "Each Boolean records a distinct one-time ODF grammar occurrence."
)]
struct Structure {
    root_seen: bool,
    body_seen: bool,
    image_seen: bool,
    frame_seen: bool,
    image_count: usize,
    body_depth: Option<usize>,
    image_depth: Option<usize>,
    frame_depth: Option<usize>,
}

impl Structure {
    fn observe(
        &mut self,
        namespace: NamespaceKind,
        element: &BytesStart<'_>,
        depth: usize,
        empty: bool,
        root: Root,
    ) -> Result<()> {
        let local = element.local_name();
        if depth == 1 {
            let expected = match root {
                Root::Flat => b"document".as_slice(),
                Root::Content => b"document-content".as_slice(),
            };
            if self.root_seen
                || namespace != NamespaceKind::Office
                || local.as_ref() != expected
                || empty
            {
                return Err(invalid("ODI requires one non-empty office document root"));
            }
            self.root_seen = true;
        } else if namespace == NamespaceKind::Office && local.as_ref() == b"body" {
            if self.body_seen || depth != 2 || empty {
                return Err(invalid("ODI requires one non-empty office:body"));
            }
            self.body_seen = true;
            self.body_depth = Some(depth);
        } else if namespace == NamespaceKind::Office && local.as_ref() == b"image" {
            if self.image_seen || self.body_depth != Some(depth - 1) || empty {
                return Err(invalid("office:image is misplaced, duplicated, or empty"));
            }
            self.image_seen = true;
            self.image_depth = Some(depth);
        } else if local.as_ref() == b"frame" {
            if namespace != NamespaceKind::Draw
                || self.frame_seen
                || self.image_depth != Some(depth - 1)
                || empty
            {
                return Err(invalid(
                    "ODI requires one non-empty draw:frame directly inside office:image",
                ));
            }
            self.frame_seen = true;
            self.frame_depth = Some(depth);
        } else if local.as_ref() == b"image" {
            if namespace != NamespaceKind::Draw || self.frame_depth != Some(depth - 1) {
                return Err(invalid(
                    "ODI requires draw:image directly inside its document frame",
                ));
            }
            self.image_count = self
                .image_count
                .checked_add(1)
                .ok_or_else(|| invalid("ODI draw:image count overflow"))?;
            if self.image_count > 1 {
                return Err(invalid(
                    "ODI document frame contains multiple draw:image elements",
                ));
            }
        }
        Ok(())
    }

    fn close(&mut self, depth: usize) {
        if self.frame_depth == Some(depth) {
            self.frame_depth = None;
        }
        if self.image_depth == Some(depth) {
            self.image_depth = None;
        }
        if self.body_depth == Some(depth) {
            self.body_depth = None;
        }
    }
}

fn scan_sites(xml: &str) -> Result<Vec<FrameSite>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut frames = Vec::<(usize, Range<usize>, FrameAttributeSites, Properties)>::new();
    let mut images = Vec::new();
    let mut image_depth = None;
    let mut binary = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|error| invalid(format!("ODI XML position exceeds usize: {error}")))?;
        let (resolved_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODI XML: {error}")))?;
        let namespace = classify(&resolved_namespace);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|error| invalid(format!("ODI XML position exceeds usize: {error}")))?;
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"frame" {
                    let z_index = attribute_value(&reader, &element, DRAW, b"z-index")?
                        .map(|value| {
                            value.parse::<u32>().map_err(|_| {
                                invalid("ODI draw:frame draw:z-index is not a non-negative integer")
                            })
                        })
                        .transpose()?;
                    frames.push((
                        depth,
                        start..end,
                        FrameAttributeSites {
                            name: attribute_qname(&reader, &element, DRAW, b"name")?,
                            style_name: attribute_qname(&reader, &element, DRAW, b"style-name")?,
                            text_style_name: attribute_qname(
                                &reader,
                                &element,
                                DRAW,
                                b"text-style-name",
                            )?,
                            layer: attribute_qname(&reader, &element, DRAW, b"layer")?,
                            z_index: attribute_qname(&reader, &element, DRAW, b"z-index")?,
                            transform: attribute_qname(&reader, &element, DRAW, b"transform")?,
                            anchor_type: attribute_qname(&reader, &element, TEXT, b"anchor-type")?,
                            x: attribute_qname(&reader, &element, SVG, b"x")?,
                            y: attribute_qname(&reader, &element, SVG, b"y")?,
                            width: attribute_qname(&reader, &element, SVG, b"width")?,
                            height: attribute_qname(&reader, &element, SVG, b"height")?,
                            relative_width: attribute_qname(
                                &reader,
                                &element,
                                STYLE,
                                b"rel-width",
                            )?,
                            relative_height: attribute_qname(
                                &reader,
                                &element,
                                STYLE,
                                b"rel-height",
                            )?,
                        },
                        Properties {
                            style_name: attribute_value(&reader, &element, DRAW, b"style-name")?,
                            text_style_name: attribute_value(
                                &reader,
                                &element,
                                DRAW,
                                b"text-style-name",
                            )?,
                            layer: attribute_value(&reader, &element, DRAW, b"layer")?,
                            z_index,
                            transform: attribute_value(&reader, &element, DRAW, b"transform")?,
                            relative_width: attribute_value(
                                &reader,
                                &element,
                                STYLE,
                                b"rel-width",
                            )?,
                            relative_height: attribute_value(
                                &reader,
                                &element,
                                STYLE,
                                b"rel-height",
                            )?,
                            copy_of: attribute_value(&reader, &element, DRAW, b"copy-of")?,
                            image_map: None,
                        },
                    ));
                } else if namespace == NamespaceKind::Draw && local.as_ref() == b"image" {
                    let (frame_tag, attributes, properties) = frames.last().map_or(
                        (None, FrameAttributeSites::default(), Properties::default()),
                        |frame| (Some(frame.1.clone()), frame.2.clone(), frame.3.clone()),
                    );
                    images.push(FrameSite {
                        frame_tag,
                        attributes,
                        properties,
                        image_tag: start..end,
                        href_attribute: attribute_qname(&reader, &element, XLINK, b"href")?,
                        binary_contents: None,
                    });
                    image_depth = Some((depth, images.len() - 1));
                } else if namespace == NamespaceKind::Office
                    && local.as_ref() == b"binary-data"
                    && let Some((_, image)) = image_depth
                {
                    binary = Some((depth, image, end));
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"image" {
                    let (frame_tag, attributes, properties) = frames.last().map_or(
                        (None, FrameAttributeSites::default(), Properties::default()),
                        |frame| (Some(frame.1.clone()), frame.2.clone(), frame.3.clone()),
                    );
                    images.push(FrameSite {
                        frame_tag,
                        attributes,
                        properties,
                        image_tag: start..end,
                        href_attribute: attribute_qname(&reader, &element, XLINK, b"href")?,
                        binary_contents: None,
                    });
                }
            },
            Event::End(_) => {
                if let Some((binary_depth, image, contents_start)) = binary
                    && binary_depth == depth
                {
                    images
                        .get_mut(image)
                        .ok_or_else(|| invalid("ODI binary image site is out of bounds"))?
                        .binary_contents = Some(contents_start..start);
                    binary = None;
                }
                if image_depth.is_some_and(|(image, _)| image == depth) {
                    image_depth = None;
                }
                if frames.last().is_some_and(|frame| frame.0 == depth) {
                    frames.pop();
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODI XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODI XML")),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(images)
}

#[derive(Clone, Copy)]
enum MapTextKind {
    Title,
    Description,
}

struct ActiveArea {
    depth: usize,
    kind: AreaKind,
    properties: AreaProperties,
    text: Option<(usize, MapTextKind, String)>,
}

#[allow(
    clippy::too_many_lines,
    reason = "One bounded streaming state machine keeps image-map placement, text, and area closure checks coherent."
)]
fn scan_image_map(xml: &str) -> Result<Option<ImageMap>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut frame_depth = None;
    let mut map_depth = None;
    let mut map_seen = false;
    let mut areas = Vec::new();
    let mut active = None::<ActiveArea>;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODI image-map XML: {error}")))?;
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                let local = element.local_name();
                if bound_to(&namespace, DRAW) && local.as_ref() == b"frame" {
                    frame_depth = Some(depth);
                } else if local.as_ref() == b"image-map" {
                    if !bound_to(&namespace, DRAW)
                        || frame_depth != depth.checked_sub(1)
                        || map_seen
                    {
                        return Err(invalid(
                            "ODI draw:image-map must be the frame's unique direct image map",
                        ));
                    }
                    map_seen = true;
                    map_depth = Some(depth);
                } else if is_area_local(local.as_ref()) {
                    if !bound_to(&namespace, DRAW) || map_depth != depth.checked_sub(1) {
                        return Err(invalid(
                            "ODI image-map area is misplaced or namespace-spoofed",
                        ));
                    }
                    ensure_map_capacity(areas.len())?;
                    active = Some(parse_area(&reader, &element, depth)?);
                } else if let Some(area) = active.as_mut()
                    && depth == area.depth + 1
                    && bound_to(&namespace, SVG)
                    && matches!(local.as_ref(), b"title" | b"desc")
                {
                    if area.text.is_some() {
                        return Err(invalid("ODI image-map accessibility text is nested"));
                    }
                    let kind = if local.as_ref() == b"title" {
                        MapTextKind::Title
                    } else {
                        MapTextKind::Description
                    };
                    area.text = Some((depth, kind, String::new()));
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if local.as_ref() == b"image-map" {
                    if !bound_to(&namespace, DRAW) || frame_depth != Some(depth) || map_seen {
                        return Err(invalid(
                            "ODI draw:image-map must be the frame's unique direct image map",
                        ));
                    }
                    map_seen = true;
                } else if is_area_local(local.as_ref()) {
                    if !bound_to(&namespace, DRAW) || map_depth != Some(depth) {
                        return Err(invalid(
                            "ODI image-map area is misplaced or namespace-spoofed",
                        ));
                    }
                    ensure_map_capacity(areas.len())?;
                    let area = parse_area(&reader, &element, depth + 1)?;
                    areas.push(Area::new(area.kind, area.properties));
                } else if let Some(area) = active.as_mut()
                    && depth == area.depth
                    && bound_to(&namespace, SVG)
                    && matches!(local.as_ref(), b"title" | b"desc")
                {
                    let target = if local.as_ref() == b"title" {
                        &mut area.properties.title
                    } else {
                        &mut area.properties.description
                    };
                    if target.replace(String::new()).is_some() {
                        return Err(invalid(
                            "ODI image-map area contains duplicate accessibility text",
                        ));
                    }
                }
            },
            Event::Text(text)
                if active
                    .as_ref()
                    .and_then(|area| area.text.as_ref())
                    .is_some() =>
            {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| invalid(format!("invalid ODI image-map text: {error}")))?;
                append_map_text(
                    &mut active
                        .as_mut()
                        .and_then(|area| area.text.as_mut())
                        .ok_or_else(|| invalid("ODI image-map text state disappeared"))?
                        .2,
                    &value,
                )?;
            },
            Event::CData(text)
                if active
                    .as_ref()
                    .and_then(|area| area.text.as_ref())
                    .is_some() =>
            {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| invalid(format!("invalid ODI image-map CDATA: {error}")))?;
                append_map_text(
                    &mut active
                        .as_mut()
                        .and_then(|area| area.text.as_mut())
                        .ok_or_else(|| invalid("ODI image-map text state disappeared"))?
                        .2,
                    &value,
                )?;
            },
            Event::GeneralRef(reference)
                if active
                    .as_ref()
                    .and_then(|area| area.text.as_ref())
                    .is_some() =>
            {
                let value = resolve_reference(&reference)?;
                append_map_text(
                    &mut active
                        .as_mut()
                        .and_then(|area| area.text.as_mut())
                        .ok_or_else(|| invalid("ODI image-map text state disappeared"))?
                        .2,
                    &value,
                )?;
            },
            Event::End(_) => {
                if let Some(area) = active.as_mut()
                    && area.text.as_ref().is_some_and(|text| text.0 == depth)
                {
                    let (_, kind, value) = area
                        .text
                        .take()
                        .ok_or_else(|| invalid("ODI image-map text state disappeared"))?;
                    let target = match kind {
                        MapTextKind::Title => &mut area.properties.title,
                        MapTextKind::Description => &mut area.properties.description,
                    };
                    if target.replace(value).is_some() {
                        return Err(invalid(
                            "ODI image-map area contains duplicate accessibility text",
                        ));
                    }
                } else if active.as_ref().is_some_and(|area| area.depth == depth) {
                    let area = active
                        .take()
                        .ok_or_else(|| invalid("ODI image-map area state disappeared"))?;
                    areas.push(Area::new(area.kind, area.properties));
                }
                if map_depth == Some(depth) {
                    map_depth = None;
                }
                if frame_depth == Some(depth) {
                    frame_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODI image-map XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODI XML")),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if active.is_some() || map_depth.is_some() || depth != 0 {
        return Err(invalid("unterminated ODI image-map XML"));
    }
    Ok(map_seen.then(|| ImageMap::new(areas)))
}

fn parse_area(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    depth: usize,
) -> Result<ActiveArea> {
    let required = |namespace, local, label| {
        attribute_value(reader, element, namespace, local)?
            .ok_or_else(|| invalid(format!("ODI image-map area requires {label}")))
    };
    let kind = match element.local_name().as_ref() {
        b"area-rectangle" => AreaKind::Rectangle {
            x: required(SVG, b"x", "svg:x")?,
            y: required(SVG, b"y", "svg:y")?,
            width: required(SVG, b"width", "svg:width")?,
            height: required(SVG, b"height", "svg:height")?,
        },
        b"area-circle" => AreaKind::Circle {
            center_x: required(SVG, b"cx", "svg:cx")?,
            center_y: required(SVG, b"cy", "svg:cy")?,
            radius: required(SVG, b"r", "svg:r")?,
        },
        b"area-polygon" => AreaKind::Polygon {
            x: required(SVG, b"x", "svg:x")?,
            y: required(SVG, b"y", "svg:y")?,
            width: required(SVG, b"width", "svg:width")?,
            height: required(SVG, b"height", "svg:height")?,
            view_box: required(SVG, b"viewBox", "svg:viewBox")?,
            points: required(DRAW, b"points", "draw:points")?,
        },
        _ => return Err(invalid("unsupported ODI image-map area")),
    };
    let no_href = match attribute_value(reader, element, DRAW, b"nohref")? {
        None => false,
        Some(value) if value == "nohref" => true,
        Some(_) => return Err(invalid("ODI image-map draw:nohref must equal 'nohref'")),
    };
    Ok(ActiveArea {
        depth,
        kind,
        properties: AreaProperties {
            href: attribute_value(reader, element, XLINK, b"href")?,
            target_frame_name: attribute_value(reader, element, OFFICE, b"target-frame-name")?,
            name: attribute_value(reader, element, OFFICE, b"name")?,
            no_href,
            link_type: attribute_value(reader, element, XLINK, b"type")?,
            show: attribute_value(reader, element, XLINK, b"show")?,
            title: None,
            description: None,
        },
        text: None,
    })
}

fn is_area_local(local: &[u8]) -> bool {
    matches!(local, b"area-rectangle" | b"area-circle" | b"area-polygon")
}

fn ensure_map_capacity(current: usize) -> Result<()> {
    if current >= MAX_MAP_AREAS {
        return Err(invalid("ODI image-map area count exceeds the limit"));
    }
    Ok(())
}

fn append_map_text(target: &mut String, value: &str) -> Result<()> {
    if target.len().saturating_add(value.len()) > MAX_MAP_TEXT_BYTES {
        return Err(invalid(
            "ODI image-map accessibility text exceeds the limit",
        ));
    }
    target.push_str(value);
    Ok(())
}

fn resolve_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        invalid(format!(
            "invalid ODI image-map character reference: {error}"
        ))
    })? {
        return Ok(character.to_string());
    }
    match reference.as_ref() {
        b"amp" => Ok("&".to_owned()),
        b"lt" => Ok("<".to_owned()),
        b"gt" => Ok(">".to_owned()),
        b"apos" => Ok("'".to_owned()),
        b"quot" => Ok("\"".to_owned()),
        _ => Err(invalid("unsupported entity in ODI image-map text")),
    }
}

fn bound_to(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn attribute_value(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut result = None;
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODI XML attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if bound_to(&resolved, namespace) && name.as_ref() == local {
            if result.is_some() {
                return Err(invalid("duplicate expanded ODI XML attribute"));
            }
            result = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| invalid(format!("invalid ODI attribute value: {error}")))?
                    .into_owned(),
            );
        }
    }
    Ok(result)
}

fn attribute_qname(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(format!("invalid ODI XML attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == namespace)
            && name.as_ref() == local
        {
            return std::str::from_utf8(attribute.key.as_ref())
                .map(str::to_owned)
                .map(Some)
                .map_err(|error| invalid(format!("ODI attribute name is not UTF-8: {error}")));
        }
    }
    Ok(None)
}

fn frame_attributes_differ(before: &Frame, after: &Frame) -> bool {
    before.name() != after.name()
        || before.style_name() != after.style_name()
        || before.text_style_name() != after.text_style_name()
        || before.layer() != after.layer()
        || before.z_index() != after.z_index()
        || before.transform() != after.transform()
        || before.anchor_type() != after.anchor_type()
        || before.x() != after.x()
        || before.y() != after.y()
        || before.width() != after.width()
        || before.height() != after.height()
        || before.relative_width() != after.relative_width()
        || before.relative_height() != after.relative_height()
}

fn rewrite_frame_tag(
    source: &str,
    sites: &FrameAttributeSites,
    before: &Frame,
    after: &Frame,
) -> Result<String> {
    let mut tag = source.to_owned();
    rewrite_changed_attribute(
        &mut tag,
        sites.name.as_deref(),
        "draw:name",
        before.name(),
        after.name(),
    )?;
    rewrite_changed_attribute(
        &mut tag,
        sites.style_name.as_deref(),
        "draw:style-name",
        before.style_name(),
        after.style_name(),
    )?;
    rewrite_changed_attribute(
        &mut tag,
        sites.text_style_name.as_deref(),
        "draw:text-style-name",
        before.text_style_name(),
        after.text_style_name(),
    )?;
    rewrite_changed_attribute(
        &mut tag,
        sites.layer.as_deref(),
        "draw:layer",
        before.layer(),
        after.layer(),
    )?;
    let before_z = before.z_index().map(|value| value.to_string());
    let after_z = after.z_index().map(|value| value.to_string());
    rewrite_changed_attribute(
        &mut tag,
        sites.z_index.as_deref(),
        "draw:z-index",
        before_z.as_deref(),
        after_z.as_deref(),
    )?;
    rewrite_changed_attribute(
        &mut tag,
        sites.transform.as_deref(),
        "draw:transform",
        before.transform(),
        after.transform(),
    )?;
    rewrite_changed_attribute(
        &mut tag,
        sites.anchor_type.as_deref(),
        "text:anchor-type",
        before.anchor_type(),
        after.anchor_type(),
    )?;
    for (site, fallback, old, new) in [
        (sites.x.as_deref(), "svg:x", before.x(), after.x()),
        (sites.y.as_deref(), "svg:y", before.y(), after.y()),
        (
            sites.width.as_deref(),
            "svg:width",
            before.width(),
            after.width(),
        ),
        (
            sites.height.as_deref(),
            "svg:height",
            before.height(),
            after.height(),
        ),
        (
            sites.relative_width.as_deref(),
            "style:rel-width",
            before.relative_width(),
            after.relative_width(),
        ),
        (
            sites.relative_height.as_deref(),
            "style:rel-height",
            before.relative_height(),
            after.relative_height(),
        ),
    ] {
        rewrite_changed_attribute(&mut tag, site, fallback, old, new)?;
    }
    Ok(tag)
}

fn rewrite_changed_attribute(
    tag: &mut String,
    existing: Option<&str>,
    fallback: &str,
    before: Option<&str>,
    after: Option<&str>,
) -> Result<()> {
    if before == after {
        return Ok(());
    }
    let escaped = after.map(|value| quick_xml::escape::escape(value).into_owned());
    *tag = match (existing, escaped) {
        (Some(name), Some(value)) => replace_attribute(tag, name, &value)?,
        (Some(name), None) => remove_attribute(tag, name)?,
        (None, Some(value)) => insert_attribute(tag, fallback, &value)?,
        (None, None) => tag.clone(),
    };
    Ok(())
}

fn rewrite_attribute(
    bytes: &[u8],
    span: &Range<usize>,
    existing: Option<&str>,
    fallback: &str,
    new_value: Option<&str>,
) -> Result<String> {
    let tag = std::str::from_utf8(
        bytes
            .get(span.clone())
            .ok_or_else(|| invalid("ODI attribute tag span is invalid"))?,
    )
    .map_err(|error| invalid(format!("ODI attribute tag is not UTF-8: {error}")))?;
    let escaped = new_value.map(|candidate| quick_xml::escape::escape(candidate).into_owned());
    match (existing, escaped) {
        (Some(name), Some(replacement)) => replace_attribute(tag, name, &replacement),
        (Some(name), None) => remove_attribute(tag, name),
        (None, Some(replacement)) => insert_attribute(tag, fallback, &replacement),
        (None, None) => Ok(tag.to_owned()),
    }
}

fn replace_attribute(tag: &str, name: &str, value: &str) -> Result<String> {
    let (_, value_span) =
        find_attribute(tag, name)?.ok_or_else(|| invalid("ODI attribute disappeared"))?;
    Ok(format!(
        "{}{}{}",
        &tag[..value_span.start],
        value,
        &tag[value_span.end..]
    ))
}

fn remove_attribute(tag: &str, name: &str) -> Result<String> {
    let (attribute, _) =
        find_attribute(tag, name)?.ok_or_else(|| invalid("ODI attribute disappeared"))?;
    Ok(format!(
        "{}{}",
        &tag[..attribute.start],
        &tag[attribute.end..]
    ))
}

fn insert_attribute(tag: &str, name: &str, value: &str) -> Result<String> {
    let position = if tag.ends_with("/>") {
        tag.len() - 2
    } else if tag.ends_with('>') {
        tag.len() - 1
    } else {
        return Err(invalid("ODI start tag has no closing delimiter"));
    };
    Ok(format!(
        "{} {}=\"{}\"{}",
        &tag[..position],
        name,
        value,
        &tag[position..]
    ))
}

fn find_attribute(tag: &str, wanted: &str) -> Result<Option<(Range<usize>, Range<usize>)>> {
    let bytes = tag.as_bytes();
    let mut cursor = 1usize;
    while cursor < bytes.len()
        && !bytes[cursor].is_ascii_whitespace()
        && bytes[cursor] != b'>'
        && bytes[cursor] != b'/'
    {
        cursor += 1;
    }
    while cursor < bytes.len() {
        let attribute_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() || bytes[cursor] == b'>' || bytes[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < bytes.len() && !bytes[cursor].is_ascii_whitespace() && bytes[cursor] != b'='
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() || bytes[cursor] != b'=' {
            return Err(invalid("ODI attribute is malformed"));
        }
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes
            .get(cursor)
            .ok_or_else(|| invalid("ODI attribute value is missing"))?;
        if quote != b'\'' && quote != b'\"' {
            return Err(invalid("ODI attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor == bytes.len() {
            return Err(invalid("ODI attribute value is unterminated"));
        }
        let value_end = cursor;
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok(Some((attribute_start..cursor, value_start..value_end)));
        }
    }
    Ok(None)
}

fn apply_edits(source: &[u8], mut edits: Vec<(Range<usize>, String)>) -> Result<Vec<u8>> {
    edits.sort_unstable_by_key(|edit| std::cmp::Reverse(edit.0.start));
    let mut output = source.to_vec();
    let mut prior = source.len();
    for (span, replacement) in edits {
        if span.start > span.end || span.end > prior || span.end > output.len() {
            return Err(invalid("overlapping or invalid ODI edit span"));
        }
        output.splice(span.clone(), replacement.bytes());
        prior = span.start;
        if output.len() > MAX_BYTES {
            return Err(invalid("flat ODI edited output exceeds the size limit"));
        }
    }
    Ok(output)
}

fn base64(bytes: &[u8]) -> String {
    const TABLE: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut output = String::with_capacity(bytes.len().div_ceil(3).saturating_mul(4));
    for chunk in bytes.chunks(3) {
        let first = chunk[0];
        let second = *chunk.get(1).unwrap_or(&0);
        let third = *chunk.get(2).unwrap_or(&0);
        output.push(TABLE[(first >> 2) as usize] as char);
        output.push(TABLE[((first & 3) << 4 | second >> 4) as usize] as char);
        output.push(if chunk.len() > 1 {
            TABLE[((second & 15) << 2 | third >> 6) as usize] as char
        } else {
            '='
        });
        output.push(if chunk.len() > 2 {
            TABLE[(third & 63) as usize] as char
        } else {
            '='
        });
    }
    output
}

fn checked_depth(depth: usize) -> Result<usize> {
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("ODI XML depth overflow"))?;
    if next_depth > MAX_DEPTH {
        return Err(invalid("ODI XML depth exceeds the limit"));
    }
    Ok(next_depth)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Draw,
    Other,
}

fn classify(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == DRAW => NamespaceKind::Draw,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
