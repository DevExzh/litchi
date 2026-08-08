//! Byte-preserving flat OpenDocument Image snapshots.

use crate::{frame::Frame, source::Source};
use litchi_core::{Error, FileFormat, Result};
use litchi_odf_common::media;
use quick_xml::{
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::sync::Arc;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const MAX_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_FRAMES: usize = 1_000_000;

#[derive(Debug)]
struct State {
    bytes: Vec<u8>,
    frames: Vec<Frame>,
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
        let xml = std::str::from_utf8(&bytes).map_err(|_| invalid("flat ODI is not UTF-8"))?;
        validate_structure(xml)?;
        let images = media::scan_flat(xml)?;
        if images.len() > MAX_FRAMES {
            return Err(invalid("flat ODI frame count exceeds the limit"));
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
        Ok(Self(Arc::new(State { bytes, frames })))
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

    /// Consumes the snapshot and returns its exact flat XML bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        Arc::try_unwrap(self.0).map_or_else(|state| state.bytes.clone(), |state| state.bytes)
    }
}

fn validate_structure(xml: &str) -> Result<()> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut body_seen = false;
    let mut image_seen = false;
    let mut body_depth = None;
    let mut image_depth = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid flat ODI XML: {error}")))?;
        let namespace = classify(&namespace);
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                observe(
                    namespace,
                    &element,
                    depth,
                    false,
                    &mut root_seen,
                    &mut body_seen,
                    &mut image_seen,
                    &mut body_depth,
                    &mut image_depth,
                )?;
            },
            Event::Empty(element) => {
                let virtual_depth = checked_depth(depth)?;
                observe(
                    namespace,
                    &element,
                    virtual_depth,
                    true,
                    &mut root_seen,
                    &mut body_seen,
                    &mut image_seen,
                    &mut body_depth,
                    &mut image_depth,
                )?;
            },
            Event::End(_) => {
                if image_depth == Some(depth) {
                    image_depth = None;
                }
                if body_depth == Some(depth) {
                    body_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("flat ODI XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in flat ODI")),
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 || !root_seen || !body_seen || !image_seen {
        return Err(invalid(
            "flat ODI requires office:document/office:body/office:image",
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
    root_seen: &mut bool,
    body_seen: &mut bool,
    image_seen: &mut bool,
    body_depth: &mut Option<usize>,
    image_depth: &mut Option<usize>,
) -> Result<()> {
    let local = element.local_name();
    if depth == 1 {
        if *root_seen
            || namespace != NamespaceKind::Office
            || local.as_ref() != b"document"
            || empty
        {
            return Err(invalid(
                "flat ODI requires one non-empty office:document root",
            ));
        }
        *root_seen = true;
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"body" {
        if *body_seen || depth != 2 || empty {
            return Err(invalid("flat ODI requires one non-empty office:body"));
        }
        *body_seen = true;
        *body_depth = Some(depth);
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"image" {
        if *image_seen || *body_depth != Some(depth - 1) {
            return Err(invalid("office:image is misplaced or duplicated"));
        }
        *image_seen = true;
        if !empty {
            *image_depth = Some(depth);
        }
    } else if namespace == NamespaceKind::Draw
        && local.as_ref() == b"image"
        && image_depth.is_none()
    {
        return Err(invalid("draw:image is outside office:image"));
    }
    Ok(())
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("flat ODI XML depth overflow"))?;
    if depth > MAX_DEPTH {
        return Err(invalid("flat ODI XML depth exceeds the limit"));
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
