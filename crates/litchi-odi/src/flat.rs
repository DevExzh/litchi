//! Byte-preserving flat OpenDocument Image snapshots and bounded frame edits.

use crate::{frame::Frame, source::Source};
use litchi_core::{Error, FileFormat, Result};
use litchi_odf_common::{compact_xml, media};
use quick_xml::{
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{ops::Range, sync::Arc};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_FRAMES: usize = 1_000_000;

#[derive(Debug)]
struct State {
    bytes: Vec<u8>,
    frames: Vec<Frame>,
    sites: Vec<FrameSite>,
}

#[derive(Clone, Debug)]
struct FrameSite {
    frame_tag: Option<Range<usize>>,
    name_attribute: Option<String>,
    image_tag: Range<usize>,
    href_attribute: Option<String>,
    binary_contents: Option<Range<usize>>,
}

/// An immutable, byte-preserving flat ODI snapshot.
#[derive(Clone, Debug)]
pub struct FlatImage(Arc<State>);

impl FlatImage {
    /// Opens a flat ODI document and inventories its inert image frames.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        if bytes.len() > MAX_BYTES {
            return Err(invalid("flat ODI exceeds the input size limit"));
        }
        if litchi_odf_common::detect::flat(&bytes) != Some(FileFormat::Odi) {
            return Err(invalid("input is not a flat ODI document"));
        }
        Ok(Self(Arc::new(parse(bytes, Root::Flat)?)))
    }

    /// Returns the inert frames in source order.
    #[must_use]
    pub fn frames(&self) -> &[Frame] {
        &self.0.frames
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
    pub fn into_bytes(self) -> Vec<u8> {
        Arc::try_unwrap(self.0).map_or_else(|state| state.bytes.clone(), |state| state.bytes)
    }
}

/// A source-bound mutable draft of flat ODI frame metadata.
pub struct FlatImageTransaction {
    source: FlatImage,
    changes: Vec<FrameChange>,
}

impl FlatImageTransaction {
    /// Stages a replacement for a frame's optional `draw:name` attribute.
    pub fn set_frame_name(&mut self, frame: usize, name: Option<String>) -> Result<()> {
        let (before_name, before_source) = {
            let current = self.frame(frame)?;
            (current.name().map(str::to_owned), current.source().clone())
        };
        if before_name.as_deref() == name.as_deref() {
            if let Some(change) = self.changes.iter_mut().find(|change| change.frame == frame) {
                change.after_name = change.before_name.clone();
            }
            self.remove_noops();
            return Ok(());
        }
        if self.site(frame)?.frame_tag.is_none() {
            return Err(invalid(
                "flat ODI image frame has no editable draw:frame owner",
            ));
        }
        if let Some(change) = self.changes.iter_mut().find(|change| change.frame == frame) {
            change.after_name = name;
        } else {
            self.changes.push(FrameChange {
                frame,
                before_name,
                after_name: name,
                before_source: before_source.clone(),
                after_source: before_source,
            });
        }
        Ok(())
    }

    /// Stages replacement of an existing linked URI or inline binary payload.
    ///
    /// Cross-kind source changes refuse rather than reconstructing unknown XML.
    pub fn set_source(&mut self, frame: usize, source: Source) -> Result<()> {
        let (before_name, before_source) = {
            let current = self.frame(frame)?;
            (current.name().map(str::to_owned), current.source().clone())
        };
        if before_source == source {
            if let Some(change) = self.changes.iter_mut().find(|change| change.frame == frame) {
                change.after_source = change.before_source.clone();
            }
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
        if let Some(change) = self.changes.iter_mut().find(|change| change.frame == frame) {
            change.after_source = source;
        } else {
            self.changes.push(FrameChange {
                frame,
                before_name: before_name.clone(),
                after_name: before_name,
                before_source,
                after_source: source,
            });
        }
        Ok(())
    }

    /// Atomically validates and publishes an immutable edited snapshot.
    pub fn commit(self) -> Result<FlatImageCommit> {
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
        let mut edits = Vec::with_capacity(self.changes.len() * 2);
        for change in &self.changes {
            let site = self.site(change.frame)?;
            if change.before_name != change.after_name {
                let tag = site
                    .frame_tag
                    .as_ref()
                    .ok_or_else(|| invalid("flat ODI frame name site disappeared"))?;
                edits.push((
                    tag.clone(),
                    rewrite_attribute(
                        self.source.as_bytes(),
                        tag,
                        site.name_attribute.as_deref(),
                        "draw:name",
                        change.after_name.as_deref(),
                    )?,
                ));
            }
            if change.before_source != change.after_source {
                match &change.after_source {
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
        let snapshot = FlatImage::from_bytes(bytes)?;
        for change in &self.changes {
            let actual = snapshot
                .frames()
                .get(change.frame)
                .ok_or_else(|| invalid("flat ODI edit lost its selected frame"))?;
            if actual.name() != change.after_name.as_deref()
                || actual.source() != &change.after_source
            {
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

    fn remove_noops(&mut self) {
        self.changes.retain(|change| {
            change.before_name != change.after_name || change.before_source != change.after_source
        });
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
    before_name: Option<String>,
    after_name: Option<String>,
    before_source: Source,
    after_source: Source,
}

impl FrameChange {
    #[must_use]
    pub fn frame(&self) -> usize {
        self.frame
    }
    #[must_use]
    pub fn before_name(&self) -> Option<&str> {
        self.before_name.as_deref()
    }
    #[must_use]
    pub fn after_name(&self) -> Option<&str> {
        self.after_name.as_deref()
    }
    #[must_use]
    pub fn before_source(&self) -> &Source {
        &self.before_source
    }
    #[must_use]
    pub fn after_source(&self) -> &Source {
        &self.after_source
    }

    fn inverse(&self) -> Self {
        Self {
            frame: self.frame,
            before_name: self.after_name.clone(),
            after_name: self.before_name.clone(),
            before_source: self.after_source.clone(),
            after_source: self.before_source.clone(),
        }
    }
}

pub(crate) fn frames_from_content(xml: &str) -> Result<Vec<Frame>> {
    if xml.len() > MAX_BYTES {
        return Err(invalid("ODI content.xml exceeds the input size limit"));
    }
    Ok(parse(xml.as_bytes().to_vec(), Root::Content)?.frames)
}

#[derive(Clone, Copy)]
enum Root {
    Flat,
    Content,
}

fn parse(bytes: Vec<u8>, root: Root) -> Result<State> {
    let xml = std::str::from_utf8(&bytes).map_err(|_| invalid("ODI XML is not UTF-8"))?;
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
    for image in images {
        let source = match image.source {
            media::Source::Inline { bytes, .. } => Source::Embedded(bytes),
            media::Source::Linked { href }
            | media::Source::PackagePart { href, .. }
            | media::Source::MissingPackagePart { href, .. } => Source::Linked(href),
            media::Source::Missing => {
                return Err(invalid(
                    "flat ODI draw:image has no losslessly modeled source",
                ));
            },
            _ => return Err(invalid("flat ODI image source is not supported")),
        };
        let mut frame = Frame::new(source);
        if let Some(name) = image.frame.and_then(|frame| frame.name) {
            frame = frame.with_name(name);
        }
        frames.push(frame);
    }
    Ok(State {
        bytes,
        frames,
        sites,
    })
}

fn validate_structure(xml: &str, root: Root) -> Result<()> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut body_seen = false;
    let mut image_seen = false;
    let mut body_depth = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODI XML: {error}")))?;
        let namespace = classify(&namespace);
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                observe(
                    namespace,
                    &element,
                    depth,
                    false,
                    root,
                    &mut root_seen,
                    &mut body_seen,
                    &mut image_seen,
                    &mut body_depth,
                )?;
            },
            Event::Empty(element) => {
                observe(
                    namespace,
                    &element,
                    checked_depth(depth)?,
                    true,
                    root,
                    &mut root_seen,
                    &mut body_seen,
                    &mut image_seen,
                    &mut body_depth,
                )?;
            },
            Event::End(_) => {
                if body_depth == Some(depth) {
                    body_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODI XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODI XML")),
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !root_seen || !body_seen || !image_seen {
        return Err(invalid(
            "ODI requires office:document/office:body/office:image",
        ));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn observe(
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
    depth: usize,
    empty: bool,
    root: Root,
    root_seen: &mut bool,
    body_seen: &mut bool,
    image_seen: &mut bool,
    body_depth: &mut Option<usize>,
) -> Result<()> {
    let local = element.local_name();
    if depth == 1 {
        let expected = match root {
            Root::Flat => b"document".as_slice(),
            Root::Content => b"document-content".as_slice(),
        };
        if *root_seen || namespace != NamespaceKind::Office || local.as_ref() != expected || empty {
            return Err(invalid("ODI requires one non-empty office document root"));
        }
        *root_seen = true;
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"body" {
        if *body_seen || depth != 2 || empty {
            return Err(invalid("ODI requires one non-empty office:body"));
        }
        *body_seen = true;
        *body_depth = Some(depth);
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"image" {
        if *image_seen || *body_depth != Some(depth - 1) {
            return Err(invalid("office:image is misplaced or duplicated"));
        }
        *image_seen = true;
    }
    Ok(())
}

fn scan_sites(xml: &str) -> Result<Vec<FrameSite>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut frames = Vec::<(usize, Range<usize>, Option<String>)>::new();
    let mut images = Vec::new();
    let mut image_depth = None;
    let mut binary = None;
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODI XML: {error}")))?;
        let namespace = classify(&namespace);
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"frame" {
                    frames.push((
                        depth,
                        start..end,
                        attribute_qname(&reader, &element, DRAW, b"name")?,
                    ));
                } else if namespace == NamespaceKind::Draw && local.as_ref() == b"image" {
                    let (frame_tag, name) = frames.last().map_or((None, None), |frame| {
                        (Some(frame.1.clone()), frame.2.clone())
                    });
                    images.push(FrameSite {
                        frame_tag,
                        name_attribute: name,
                        image_tag: start..end,
                        href_attribute: attribute_qname(&reader, &element, XLINK, b"href")?,
                        binary_contents: None,
                    });
                    image_depth = Some((depth, images.len() - 1));
                } else if namespace == NamespaceKind::Office && local.as_ref() == b"binary-data" {
                    if let Some((_, image)) = image_depth {
                        binary = Some((depth, image, end));
                    }
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"image" {
                    let (frame_tag, name) = frames.last().map_or((None, None), |frame| {
                        (Some(frame.1.clone()), frame.2.clone())
                    });
                    images.push(FrameSite {
                        frame_tag,
                        name_attribute: name,
                        image_tag: start..end,
                        href_attribute: attribute_qname(&reader, &element, XLINK, b"href")?,
                        binary_contents: None,
                    });
                }
            },
            Event::End(_) => {
                if let Some((binary_depth, image, contents_start)) = binary {
                    if binary_depth == depth {
                        images
                            .get_mut(image)
                            .ok_or_else(|| invalid("ODI binary image site is out of bounds"))?
                            .binary_contents = Some(contents_start..start);
                        binary = None;
                    }
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
            _ => {},
        }
    }
    Ok(images)
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
                .map_err(|_| invalid("ODI attribute name is not UTF-8"));
        }
    }
    Ok(None)
}

fn rewrite_attribute(
    bytes: &[u8],
    span: &Range<usize>,
    existing: Option<&str>,
    fallback: &str,
    value: Option<&str>,
) -> Result<String> {
    let tag = std::str::from_utf8(
        bytes
            .get(span.clone())
            .ok_or_else(|| invalid("ODI attribute tag span is invalid"))?,
    )
    .map_err(|_| invalid("ODI attribute tag is not UTF-8"))?;
    let escaped = value.map(|value| quick_xml::escape::escape(value).into_owned());
    match (existing, escaped) {
        (Some(name), Some(value)) => replace_attribute(tag, name, &value),
        (Some(name), None) => remove_attribute(tag, name),
        (None, Some(value)) => insert_attribute(tag, fallback, &value),
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
    edits.sort_unstable_by(|left, right| right.0.start.cmp(&left.0.start));
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
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("ODI XML depth overflow"))?;
    if depth > MAX_DEPTH {
        return Err(invalid("ODI XML depth exceeds the limit"));
    }
    Ok(depth)
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
        _ => NamespaceKind::Other,
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
