//! Bounded batches of structural edits on retained XML.

use super::{
    MAX_EDITS, MAX_OWNED_XML_BYTES, OwnedXmlPart, invalid_source, owned_offset, validate_source_xml,
};
use crate::{OpcError, ReadLimits, Result};
use litchi_core::xml::ReaderOrigin;
use quick_xml::{events::Event, reader::NsReader};
use std::{fmt, ops::Range, sync::Arc};

/// A structural edit addressed by an exact source opening tag.
pub struct OwnedElementUpdate<'a> {
    pub start_tag: Range<usize>,
    pub edit: OwnedElementEdit<'a>,
}

/// Insertions and replacements must contain one compact, self-contained element.
pub enum OwnedElementEdit<'a> {
    InsertBefore(&'a [u8]),
    AppendChild(&'a [u8]),
    Replace(&'a [u8]),
    Remove,
}

impl fmt::Debug for OwnedElementEdit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InsertBefore(bytes) => f.debug_tuple("InsertBefore").field(&bytes.len()).finish(),
            Self::AppendChild(bytes) => f.debug_tuple("AppendChild").field(&bytes.len()).finish(),
            Self::Replace(bytes) => f.debug_tuple("Replace").field(&bytes.len()).finish(),
            Self::Remove => f.write_str("Remove"),
        }
    }
}

impl fmt::Debug for OwnedElementUpdate<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("OwnedElementUpdate")
            .field("start_tag", &self.start_tag)
            .field("edit", &self.edit)
            .finish()
    }
}

struct Bounds {
    full: Range<usize>,
    close: usize,
    empty: bool,
    name: Range<usize>,
}

struct Target {
    tag: Range<usize>,
    updates: Range<usize>,
    bounds: Option<Bounds>,
}

enum Output<'a> {
    Raw(&'a [u8]),
    EmptyAppend {
        updates: Range<usize>,
        name: Range<usize>,
    },
}

struct Splice<'a> {
    range: Range<usize>,
    output: Output<'a>,
}

impl OwnedXmlPart {
    /// Apply ordered structural updates with one source scan and one result
    /// allocation. Targets may repeat; overlapping destructive edits fail.
    /// Multiple children appended to an empty parent share one expansion.
    /// New fragments pass the compact authoring audit and XML validation.
    pub fn update_elements(
        &self,
        updates: &[OwnedElementUpdate<'_>],
        max_output_bytes: usize,
    ) -> Result<Self> {
        let maximum = max_output_bytes.min(MAX_OWNED_XML_BYTES);
        if updates.len() > MAX_EDITS {
            return Err(invalid_source("too many owned XML element updates"));
        }
        let mut targets: Vec<Target> = Vec::new();
        targets
            .try_reserve_exact(updates.len())
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML element targets",
                source,
            })?;
        for (index, update) in updates.iter().enumerate() {
            let tag = &update.start_tag;
            if tag.start >= tag.end || tag.end > self.bytes.len() {
                return Err(invalid_source("invalid XML opening-tag range"));
            }
            if let Some(last) = targets.last_mut() {
                if last.tag == *tag {
                    last.updates.end = index + 1;
                    continue;
                }
                if tag.start < last.tag.end {
                    return Err(invalid_source("unordered XML element targets"));
                }
            }
            targets.push(Target {
                tag: tag.clone(),
                updates: index..index + 1,
                bounds: None,
            });
        }
        let mut reader = NsReader::from_reader(self.bytes.as_slice());
        let origin = ReaderOrigin::of(self.bytes.as_slice());
        let mut stack = Vec::new();
        let mut next = 0usize;
        loop {
            let start = owned_offset(&reader, origin)?;
            let event = reader
                .read_event()
                .map_err(|error| invalid_source(error.to_string()))?;
            let end = owned_offset(&reader, origin)?;
            if next < targets.len() && start > targets[next].tag.start {
                return Err(invalid_source(
                    "element update does not identify an opening tag",
                ));
            }
            let empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    let index = if next < targets.len() && targets[next].tag == (start..end) {
                        let name = element.name();
                        let at = (name.as_ref().as_ptr() as usize)
                            .checked_sub(self.bytes.as_ptr() as usize)
                            .ok_or_else(|| invalid_source("element name is not source-backed"))?;
                        targets[next].bounds = Some(Bounds {
                            full: start..end,
                            close: end - 2,
                            empty,
                            name: at..at + name.as_ref().len(),
                        });
                        let index = next;
                        next += 1;
                        Some(index)
                    } else {
                        None
                    };
                    if !empty {
                        stack.push(index);
                    }
                },
                Event::End(_) => {
                    if let Some(index) = stack
                        .pop()
                        .ok_or_else(|| invalid_source("unbalanced XML source"))?
                    {
                        let bounds = targets[index]
                            .bounds
                            .as_mut()
                            .ok_or_else(|| invalid_source("missing element bounds"))?;
                        bounds.full.end = end;
                        bounds.close = start;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if next != targets.len() || !stack.is_empty() {
            return Err(invalid_source(
                "element update does not identify a complete element",
            ));
        }
        let mut splices = Vec::new();
        splices
            .try_reserve_exact(updates.len())
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML element spans",
                source,
            })?;
        let mut authored_bytes = 0usize;
        for target in targets {
            let bounds = target
                .bounds
                .ok_or_else(|| invalid_source("missing element bounds"))?;
            let mut appended_empty = false;
            for update in &updates[target.updates.clone()] {
                if matches!(&update.edit, OwnedElementEdit::Replace(bytes) if *bytes == &self.bytes[bounds.full.clone()])
                {
                    continue;
                }
                let fragment = match &update.edit {
                    OwnedElementEdit::InsertBefore(bytes)
                    | OwnedElementEdit::AppendChild(bytes)
                    | OwnedElementEdit::Replace(bytes) => *bytes,
                    OwnedElementEdit::Remove => &[],
                };
                authored_bytes = authored_bytes
                    .checked_add(fragment.len())
                    .ok_or_else(|| invalid_source("authored element size overflow"))?;
                if authored_bytes > maximum {
                    return Err(invalid_source("authored elements exceed output limit"));
                }
                if !matches!(update.edit, OwnedElementEdit::Remove) {
                    let _report = xml_minifier::audit::verify_authored(
                        fragment,
                        xml_minifier::audit::Limits::default(),
                    )
                    .map_err(|source| OpcError::XmlPublication {
                        part: self.name.to_string(),
                        source,
                    })?;
                    validate_source_xml(&self.name, fragment, ReadLimits::default(), None)?;
                }
                let (range, output) = match update.edit {
                    OwnedElementEdit::InsertBefore(_) => {
                        (bounds.full.start..bounds.full.start, Output::Raw(fragment))
                    },
                    OwnedElementEdit::AppendChild(_) if bounds.empty => {
                        if appended_empty {
                            continue;
                        }
                        appended_empty = true;
                        (
                            bounds.close..bounds.full.end,
                            Output::EmptyAppend {
                                updates: target.updates.clone(),
                                name: bounds.name.clone(),
                            },
                        )
                    },
                    OwnedElementEdit::AppendChild(_) => {
                        (bounds.close..bounds.close, Output::Raw(fragment))
                    },
                    OwnedElementEdit::Replace(_) | OwnedElementEdit::Remove => {
                        (bounds.full.clone(), Output::Raw(fragment))
                    },
                };
                splices.push(Splice { range, output });
            }
        }
        // Parent appends can follow edits to descendants. Empty insertions at a
        // removed element's start precede the removal, preserving their order.
        splices.sort_by(|a, b| {
            a.range
                .start
                .cmp(&b.range.start)
                .then(a.range.end.cmp(&b.range.end))
        });
        let mut size = self.bytes.len();
        let mut cursor = 0usize;
        for splice in &splices {
            if splice.range.start < cursor {
                return Err(invalid_source("overlapping XML element edits"));
            }
            cursor = splice.range.end;
            let added = match &splice.output {
                Output::Raw(bytes) => bytes.len(),
                Output::EmptyAppend {
                    updates: group,
                    name,
                } => {
                    let mut added = name.len() + 4;
                    for update in &updates[group.clone()] {
                        if let OwnedElementEdit::AppendChild(bytes) = update.edit {
                            added = added
                                .checked_add(bytes.len())
                                .ok_or_else(|| invalid_source("XML append size overflow"))?;
                        }
                    }
                    added
                },
            };
            size = size
                .checked_sub(splice.range.len())
                .and_then(|size| size.checked_add(added))
                .ok_or_else(|| invalid_source("XML element output size overflow"))?;
        }
        if size > maximum {
            return Err(invalid_source("owned XML output exceeds limit"));
        }
        if splices.is_empty() {
            return Ok(self.clone());
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(size)
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML element output",
                source,
            })?;
        cursor = 0;
        for splice in splices {
            bytes.extend_from_slice(&self.bytes[cursor..splice.range.start]);
            match splice.output {
                Output::Raw(value) => bytes.extend_from_slice(value),
                Output::EmptyAppend {
                    updates: group,
                    name,
                } => {
                    bytes.push(b'>');
                    for update in &updates[group] {
                        if let OwnedElementEdit::AppendChild(value) = update.edit {
                            bytes.extend_from_slice(value);
                        }
                    }
                    bytes.extend_from_slice(b"</");
                    bytes.extend_from_slice(&self.bytes[name]);
                    bytes.push(b'>');
                },
            }
            cursor = splice.range.end;
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
