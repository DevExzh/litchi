//! Picture image replacement and shape descriptions (record 0333).
//!
//! [`Edit::set_picture_image`] points a picture's `a:blip@r:embed` at a new
//! image part and drops the old relationship and part when nothing else
//! uses them; [`Edit::set_shape_description`] sets or removes the
//! `p:cNvPr@descr` alternative text of any shape-tree element. Both are
//! lexical: only the changed attribute value and the relationship and part
//! entries change.

use std::collections::HashSet;
use std::ops::Range;

use litchi_opc::{BlobPart, PackURI};
use quick_xml::Reader;
use quick_xml::events::Event;

use super::{
    Edit, available_transfer_name, conformance_image_relationship, invalid,
    remove_unreferenced_dependencies, shape_relationship_ids,
};
use crate::{Error, Result};

/// The largest description written, in bytes.
const MAX_DESCRIPTION_BYTES: usize = 64 * 1024;

/// The start-tag span of the first element named `local` in `fragment`.
pub(super) fn first_start_tag(fragment: &[u8], local: &[u8]) -> Result<Option<Range<usize>>> {
    let mut reader = Reader::from_reader(fragment);
    loop {
        let before = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let event = reader
            .read_event()
            .map_err(|_| invalid("opened-presentation shape XML is malformed"))?;
        let after = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == local =>
            {
                return Ok(Some(before..after));
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

/// The first attribute whose local name is `local` (any prefix) of a raw
/// start tag: its value span, its qualified name and its whole span.
pub(super) fn attribute_value(
    tag: &[u8],
    local: &[u8],
) -> Option<(Range<usize>, Vec<u8>, Range<usize>)> {
    let mut at = 1;
    while at < tag.len() && !tag[at].is_ascii_whitespace() && tag[at] != b'>' && tag[at] != b'/' {
        at += 1;
    }
    loop {
        while at < tag.len() && tag[at].is_ascii_whitespace() {
            at += 1;
        }
        let start = at;
        while at < tag.len() && tag[at] != b'=' && !tag[at].is_ascii_whitespace() {
            if tag[at] == b'>' || tag[at] == b'/' {
                return None;
            }
            at += 1;
        }
        let key = tag[start..at].to_vec();
        while at < tag.len() && tag[at].is_ascii_whitespace() {
            at += 1;
        }
        if tag.get(at) != Some(&b'=') {
            return None;
        }
        at += 1;
        while at < tag.len() && tag[at].is_ascii_whitespace() {
            at += 1;
        }
        let quote = *tag.get(at)?;
        if quote != b'"' && quote != b'\'' {
            return None;
        }
        at += 1;
        let value = at;
        while at < tag.len() && tag[at] != quote {
            at += 1;
        }
        if at >= tag.len() {
            return None;
        }
        let name = key
            .iter()
            .position(|byte| *byte == b':')
            .map_or(&key[..], |colon| &key[colon + 1..]);
        if name == local {
            return Some((value..at, key, start..at + 1));
        }
        at += 1;
    }
}

fn escape_attribute(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\t' => out.push_str("&#9;"),
            '\n' => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            _ => out.push(ch),
        }
    }
    out
}

impl Edit {
    /// Replace the image a picture shows: the picture's `a:blip@r:embed`
    /// points at a new image part holding `resource`, and the old
    /// relationship and image part are removed when nothing else uses them.
    /// Returns whether the slide changed; the picture's current bytes are an
    /// exact no-op. Media pictures and SVG pictures are refused.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a shape that is not a
    /// picture, a picture without an embedded image, invalid resource
    /// metadata, empty or oversized data, or markup-compatibility content. A
    /// failure leaves the edit unchanged.
    pub fn set_picture_image<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        resource: &crate::media_parts::Resource,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        if resource.data.is_empty() {
            return Err(invalid(
                "opened-presentation picture resource cannot be empty",
            ));
        }
        let max_patch_bytes = self.source_ref()?.limits.max_patch_bytes();
        if resource.data.len() > max_patch_bytes {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::PatchBytes,
                observed: resource.data.len(),
                maximum: max_patch_bytes,
            });
        }
        if !resource.content_type.starts_with("image/") {
            return Err(invalid(
                "opened-presentation picture resource must use an image content type",
            ));
        }
        let preferred = PackURI::new(&resource.part_name)
            .map_err(|_error| invalid("opened-presentation resource part name is invalid"))?;
        if !preferred.as_str().starts_with("/ppt/media/") {
            return Err(invalid(
                "opened-presentation picture resource must use a /ppt/media/ part name",
            ));
        }
        let key = shape.into();
        let owner = self.working.get_part(&selected.part_name)?;
        let xml = owner.blob();
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(xml, shape_limits)?;
        if scene.is_rewritten() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "set_picture_image",
                "picture edits do not select markup-compatibility branches",
            ));
        }
        let picture = scene.shape(key)?;
        if !matches!(picture, crate::shape::Shape::Picture(_)) {
            return Err(invalid("the selected shape is not a picture"));
        }
        if crate::media_parts::summary(picture)?.is_some() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "set_picture_image",
                "a media picture's poster is not replaced as a picture image",
            ));
        }
        let span = picture.common().span()?.range(xml.len())?;
        let fragment = &xml[span.clone()];
        if fragment.windows(7).any(|window| window == b"svgBlip") {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "set_picture_image",
                "an SVG picture's image is not replaced",
            ));
        }
        let tag = first_start_tag(fragment, b"blip")?
            .ok_or_else(|| invalid("the picture has no a:blip"))?;
        let (value, _, _) = attribute_value(&fragment[tag.clone()], b"embed")
            .ok_or_else(|| invalid("the picture's image is not embedded"))?;
        let value = tag.start + value.start..tag.start + value.end;
        let old_id = std::str::from_utf8(&fragment[value.clone()])
            .map_err(|_| invalid("the picture's relationship ID is not UTF-8"))?
            .to_owned();
        let old_target = owner
            .rels()
            .get(&old_id)
            .filter(|relationship| !relationship.is_external())
            .map(litchi_opc::Relationship::target_partname)
            .transpose()?
            .ok_or_else(|| invalid("the picture's image relationship is missing"))?;
        if self
            .working
            .get_part(&old_target)
            .is_ok_and(|part| part.blob() == resource.data.as_slice())
        {
            return Ok(false);
        }
        let conformance = crate::media_parts::document_conformance_with_limits(
            xml,
            self.semantic_limits()
                .media_parts()
                .constrained_by_source(self.working.read_limits()),
        )?;
        let reserved: HashSet<_> = self
            .working
            .iter_parts()
            .map(|part| part.partname().clone())
            .collect();
        let part_name = available_transfer_name(&preferred, &reserved)?;
        let mut candidate = self.working.clone();
        candidate.try_add_part(Box::new(BlobPart::new(
            part_name.clone(),
            resource.content_type.clone(),
            resource.data.as_slice().to_vec(),
        )))?;
        let target = part_name.relative_ref(selected.part_name.base_uri());
        let new_id = candidate
            .get_part_mut(&selected.part_name)?
            .relate_to(&target, conformance_image_relationship(conformance));
        let absolute = span.start + value.start..span.start + value.end;
        let mut next = Vec::with_capacity(xml.len() + new_id.len());
        next.extend_from_slice(&xml[..absolute.start]);
        next.extend_from_slice(new_id.as_bytes());
        next.extend_from_slice(&xml[absolute.end..]);
        let staged = crate::shape::Scene::read(&next, shape_limits)?;
        if staged.len() != scene.len()
            || scene
                .iter()
                .zip(staged.iter())
                .any(|(left, right)| left.id() != right.id())
        {
            return Err(invalid("the picture image edit did not round-trip"));
        }
        let retained =
            shape_relationship_ids(&next, candidate.get_part(&selected.part_name)?.rels())?;
        let mut roots = Vec::new();
        {
            let owner = candidate.get_part_mut(&selected.part_name)?;
            owner.set_blob(next);
            if !retained.contains(&old_id)
                && let Some(relationship) = owner.rels_mut().remove(&old_id)
                && !relationship.is_external()
            {
                roots.push(relationship.target_partname()?);
            }
        }
        remove_unreferenced_dependencies(&mut candidate, roots)?;
        self.working = candidate;
        self.note_scene_read(&selected.part_name);
        Ok(true)
    }

    /// Set (`Some`) or remove (`None`) the alternative text (`p:cNvPr@descr`)
    /// of a shape-tree element of a slide. Returns whether the slide
    /// changed; an equal description is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a description with a
    /// character XML cannot carry or over 64 KiB, or markup-compatibility
    /// content.
    pub fn set_shape_description<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        description: Option<&str>,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        if let Some(description) = description {
            if description.len() > MAX_DESCRIPTION_BYTES {
                return Err(invalid("opened-presentation description is too long"));
            }
            if !description.chars().all(super::is_xml_char) {
                return Err(invalid(
                    "opened-presentation description contains an invalid XML character",
                ));
            }
        }
        let key = shape.into();
        let owner = self.working.get_part(&selected.part_name)?;
        let xml = owner.blob();
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(xml, shape_limits)?;
        if scene.is_rewritten() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "set_shape_description",
                "shape edits do not select markup-compatibility branches",
            ));
        }
        let target = scene.shape(key)?;
        if target.description()?.as_deref() == description {
            return Ok(false);
        }
        let span = target.common().span()?.range(xml.len())?;
        let tag = first_start_tag(&xml[span.clone()], b"cNvPr")?
            .ok_or_else(|| invalid("the shape has no p:cNvPr"))?;
        let tag = span.start + tag.start..span.start + tag.end;
        let raw = &xml[tag.clone()];
        let current =
            attribute_value(raw, b"descr").filter(|(_, key, _)| key.as_slice() == b"descr");
        let rewritten = match (current, description) {
            (Some((value, _, _)), Some(description)) => {
                let mut out = raw[..value.start].to_vec();
                out.extend_from_slice(escape_attribute(description).as_bytes());
                out.extend_from_slice(&raw[value.end..]);
                out
            },
            (Some((_, _, whole)), None) => {
                // Drop ` descr="…"` with the whitespace before it.
                let mut start = whole.start;
                while start > 0 && raw[start - 1].is_ascii_whitespace() {
                    start -= 1;
                }
                let mut out = raw[..start].to_vec();
                out.extend_from_slice(&raw[whole.end..]);
                out
            },
            (None, Some(description)) => {
                let close = if raw.ends_with(b"/>") {
                    raw.len() - 2
                } else {
                    raw.len() - 1
                };
                let mut out = raw[..close].to_vec();
                out.extend_from_slice(b" descr=\"");
                out.extend_from_slice(escape_attribute(description).as_bytes());
                out.push(b'"');
                out.extend_from_slice(&raw[close..]);
                out
            },
            (None, None) => return Ok(false),
        };
        let mut next = Vec::with_capacity(xml.len() + rewritten.len());
        next.extend_from_slice(&xml[..tag.start]);
        next.extend_from_slice(&rewritten);
        next.extend_from_slice(&xml[tag.end..]);
        let staged = crate::shape::Scene::read(&next, shape_limits)?;
        if staged.len() != scene.len()
            || staged.shape(key)?.description()?.as_deref() != description
            || scene
                .iter()
                .zip(staged.iter())
                .any(|(left, right)| left.id() != right.id())
        {
            return Err(invalid("the shape description edit did not round-trip"));
        }
        self.working
            .get_part_mut(&selected.part_name)?
            .set_blob(next);
        self.note_scene_read(&selected.part_name);
        Ok(true)
    }

    /// Rename a shape-tree element of a slide (`p:cNvPr@name`, record
    /// 0333). Returns whether the slide changed; an equal name is an exact
    /// no-op.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, an empty name, a name
    /// with a character XML cannot carry or over 64 KiB, or
    /// markup-compatibility content.
    pub fn set_shape_name<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        name: &str,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        if name.is_empty()
            || name.len() > MAX_DESCRIPTION_BYTES
            || !name.chars().all(super::is_xml_char)
        {
            return Err(invalid("opened-presentation shape name is invalid"));
        }
        let key = shape.into();
        let owner = self.working.get_part(&selected.part_name)?;
        let xml = owner.blob();
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(xml, shape_limits)?;
        if scene.is_rewritten() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "set_shape_name",
                "shape edits do not select markup-compatibility branches",
            ));
        }
        let target = scene.shape(key)?;
        let span = target.common().span()?.range(xml.len())?;
        let tag = first_start_tag(&xml[span.clone()], b"cNvPr")?
            .ok_or_else(|| invalid("the shape has no p:cNvPr"))?;
        let tag = span.start + tag.start..span.start + tag.end;
        let raw = &xml[tag.clone()];
        let escaped = escape_attribute(name);
        let rewritten =
            match attribute_value(raw, b"name").filter(|(_, key, _)| key.as_slice() == b"name") {
                Some((value, _, _)) => {
                    if &raw[value.clone()] == escaped.as_bytes() {
                        return Ok(false);
                    }
                    let mut out = raw[..value.start].to_vec();
                    out.extend_from_slice(escaped.as_bytes());
                    out.extend_from_slice(&raw[value.end..]);
                    out
                },
                None => {
                    let close = if raw.ends_with(b"/>") {
                        raw.len() - 2
                    } else {
                        raw.len() - 1
                    };
                    let mut out = raw[..close].to_vec();
                    out.extend_from_slice(b" name=\"");
                    out.extend_from_slice(escaped.as_bytes());
                    out.push(b'"');
                    out.extend_from_slice(&raw[close..]);
                    out
                },
            };
        let mut next = Vec::with_capacity(xml.len() + rewritten.len());
        next.extend_from_slice(&xml[..tag.start]);
        next.extend_from_slice(&rewritten);
        next.extend_from_slice(&xml[tag.end..]);
        let staged = crate::shape::Scene::read(&next, shape_limits)?;
        if staged.len() != scene.len()
            || scene
                .iter()
                .zip(staged.iter())
                .any(|(left, right)| left.id() != right.id())
        {
            return Err(invalid("the shape name edit did not round-trip"));
        }
        self.working
            .get_part_mut(&selected.part_name)?
            .set_blob(next);
        self.note_scene_read(&selected.part_name);
        Ok(true)
    }
}
