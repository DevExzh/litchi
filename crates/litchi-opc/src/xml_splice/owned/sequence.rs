//! Same-parent sequence publication without transplanting inherited context.

use super::{MAX_EDITS, MAX_OWNED_XML_BYTES, OwnedXmlPart, invalid_source, validate_source_xml};
use crate::{OpcError, ReadLimits, Result};
use quick_xml::{events::Event, reader::NsReader};
use std::{fmt, ops::Range, sync::Arc};

/// One element in a replacement child sequence.
pub enum OwnedChildElement<'a> {
    /// Reuse a complete child from the selected source sequence, by position.
    /// Repeated positions copy the source element under its original parent.
    Retained(usize),
    /// Insert one compact, independently valid XML element.
    Authored(&'a [u8]),
}

impl fmt::Debug for OwnedChildElement<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Retained(index) => f.debug_tuple("Retained").field(index).finish(),
            Self::Authored(bytes) => f.debug_tuple("Authored").field(&bytes.len()).finish(),
        }
    }
}

impl OwnedXmlPart {
    /// Replace selected direct children of one parent with a new sequence.
    ///
    /// Opening-tag ranges identify ordered source children. Complete retained
    /// subtrees can move or repeat only beneath this same parent, preserving
    /// their inherited XML and namespace context. Unselected source spans stay
    /// in their original slots; excess new elements follow the last selected
    /// slot, or append to the parent when no children were selected.
    ///
    /// All sizes are checked before result allocation. New fragments pass the
    /// compact authoring audit; the assembled document is fully XML-validated.
    pub fn replace_child_sequence(
        &self,
        parent_tag: Range<usize>,
        selected_tags: &[Range<usize>],
        children: &[OwnedChildElement<'_>],
        max_output_bytes: usize,
    ) -> Result<Self> {
        let maximum = max_output_bytes.min(MAX_OWNED_XML_BYTES);
        if selected_tags.len() > MAX_EDITS || children.len() > MAX_EDITS {
            return Err(invalid_source("too many owned XML child elements"));
        }
        if parent_tag.start >= parent_tag.end || parent_tag.end > self.bytes.len() {
            return Err(invalid_source("invalid XML parent opening-tag range"));
        }
        let mut previous = parent_tag.end;
        for tag in selected_tags {
            if tag.start < previous || tag.start >= tag.end || tag.end > self.bytes.len() {
                return Err(invalid_source(
                    "invalid or unordered XML child opening tags",
                ));
            }
            previous = tag.end;
        }
        let mut bounds = Vec::new();
        bounds
            .try_reserve_exact(selected_tags.len())
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML child bounds",
                source,
            })?;
        let mut reader = NsReader::from_reader(self.bytes.as_slice());
        let mut depth = 0usize;
        let mut parent_depth = None;
        let mut parent_name = None;
        let mut parent_close = None;
        let mut empty_parent = false;
        let mut active_child: Option<(usize, usize)> = None;
        let mut next = 0usize;
        loop {
            let start = reader.buffer_position() as usize;
            let event = reader
                .read_event()
                .map_err(|error| invalid_source(error.to_string()))?;
            let end = reader.buffer_position() as usize;
            let paired = matches!(&event, Event::Start(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if (start..end) == parent_tag {
                        parent_depth = Some(depth);
                        let name = element.name();
                        let at = (name.as_ref().as_ptr() as usize)
                            .checked_sub(self.bytes.as_ptr() as usize)
                            .ok_or_else(|| {
                                invalid_source("XML parent name is not source-backed")
                            })?;
                        parent_name = Some(at..at + name.as_ref().len());
                        if !paired {
                            empty_parent = true;
                            parent_close = Some(end - 2);
                        }
                    } else if next < selected_tags.len() && selected_tags[next] == (start..end) {
                        if empty_parent
                            || parent_close.is_some()
                            || parent_depth != depth.checked_sub(1)
                        {
                            return Err(invalid_source(
                                "selected XML element is not a direct child",
                            ));
                        }
                        bounds.push(start..end);
                        if paired {
                            active_child = Some((next, depth));
                        }
                        next += 1;
                    }
                    if paired {
                        depth += 1;
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid_source("unbalanced XML source"))?;
                    if let Some((index, child_depth)) = active_child {
                        if child_depth == depth {
                            bounds[index].end = end;
                            active_child = None;
                        }
                    }
                    if parent_depth == Some(depth) && parent_close.is_none() {
                        parent_close = Some(start);
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        let close =
            parent_close.ok_or_else(|| invalid_source("XML parent opening tag was not found"))?;
        let name =
            parent_name.ok_or_else(|| invalid_source("XML parent opening tag was not found"))?;
        if next != selected_tags.len() || active_child.is_some() {
            return Err(invalid_source("XML child opening tag was not found"));
        }
        let mut size = self.bytes.len();
        for range in &bounds {
            size = size
                .checked_sub(range.len())
                .ok_or_else(|| invalid_source("overlapping XML child ranges"))?;
        }
        let mut authored = 0usize;
        for child in children {
            let added = match child {
                OwnedChildElement::Retained(index) => bounds
                    .get(*index)
                    .ok_or_else(|| invalid_source("retained XML child position is out of bounds"))?
                    .len(),
                OwnedChildElement::Authored(bytes) => {
                    authored = authored
                        .checked_add(bytes.len())
                        .ok_or_else(|| invalid_source("authored XML child size overflow"))?;
                    if authored > maximum {
                        return Err(invalid_source("authored XML children exceed output limit"));
                    }
                    let _report = xml_minifier::audit::verify_authored(
                        bytes,
                        xml_minifier::audit::Limits::default(),
                    )
                    .map_err(|source| OpcError::XmlPublication {
                        part: self.name.to_string(),
                        source,
                    })?;
                    validate_source_xml(&self.name, bytes, ReadLimits::default(), None)?;
                    bytes.len()
                },
            };
            size = size
                .checked_add(added)
                .ok_or_else(|| invalid_source("XML child output size overflow"))?;
        }
        if empty_parent && !children.is_empty() {
            size = size
                .checked_add(name.len() + 2)
                .ok_or_else(|| invalid_source("XML parent expansion size overflow"))?;
        }
        if size > maximum {
            return Err(invalid_source("owned XML child output exceeds limit"));
        }
        if children.len() == bounds.len()
            && children.iter().enumerate().all(
                |(index, child)| matches!(child, OwnedChildElement::Retained(old) if *old == index),
            )
        {
            return Ok(self.clone());
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(size)
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML child output",
                source,
            })?;
        let append = |output: &mut Vec<u8>, child: &OwnedChildElement<'_>| {
            output.extend_from_slice(match child {
                OwnedChildElement::Retained(index) => &self.bytes[bounds[*index].clone()],
                OwnedChildElement::Authored(bytes) => bytes,
            });
        };
        let mut cursor = 0usize;
        for (index, range) in bounds.iter().enumerate() {
            bytes.extend_from_slice(&self.bytes[cursor..range.start]);
            if let Some(child) = children.get(index) {
                append(&mut bytes, child);
            }
            cursor = range.end;
        }
        if bounds.is_empty() {
            bytes.extend_from_slice(&self.bytes[..close]);
            cursor = close;
            if empty_parent {
                bytes.push(b'>');
                cursor += 2;
            }
        }
        for child in children.iter().skip(bounds.len()) {
            append(&mut bytes, child);
        }
        if empty_parent {
            bytes.extend_from_slice(b"</");
            bytes.extend_from_slice(&self.bytes[name]);
            bytes.push(b'>');
        }
        bytes.extend_from_slice(&self.bytes[cursor..]);
        validate_source_xml(&self.name, &bytes, ReadLimits::default(), None)?;
        Ok(Self {
            name: self.name.clone(),
            content_type: self.content_type.clone(),
            bytes: Arc::new(bytes),
        })
    }
}
