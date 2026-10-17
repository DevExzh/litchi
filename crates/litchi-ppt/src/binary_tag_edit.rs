//! Record-level edits of one programmable binary tag's data blob (record
//! 0323).
//!
//! `PowerPoint` 2000 and later keep version-specific records in the
//! `BinaryTagDataBlob` of a `___PPT9`/`___PPT10` programmable tag
//! ([MS-PPT] 2.4.22.?): a slide's tags are a direct `SlideProgTagsContainer`
//! child, the document's live in its `DocInfoListContainer`. An edit
//! replaces the blob's records and re-frames only the blob, its tag, the
//! tag container and the owner; every other record keeps its bytes. A
//! missing tag (and tag container) is created, empty, around the new
//! records.

use crate::consts::RecordType;
use crate::package::Error;
use crate::record_wire as wire;
use crate::slide_order::{Error as EditError, Refusal, Result};

const PROG_TAGS: u16 = RecordType::ProgTags as u16;
const PROG_BINARY_TAG: u16 = RecordType::ProgBinaryTag as u16;
const BINARY_TAG_DATA: u16 = RecordType::BinaryTagData as u16;
const CSTRING: u16 = RecordType::CString as u16;

/// One owner's binary tag and the records of its data blob.
pub(crate) struct Blob<'a> {
    name: &'static str,
    owner_children: Vec<&'a [u8]>,
    /// Index of the programmable-tag container among the owner's children.
    prog_tags: Option<usize>,
    /// Children of the programmable-tag container.
    tags: Vec<&'a [u8]>,
    /// Index of the named tag among `tags`.
    tag: Option<usize>,
    /// Children of the named tag.
    tag_children: Vec<&'a [u8]>,
    /// Index of the `BinaryTagData` child among `tag_children`.
    blob: Option<usize>,
    /// Records of the blob.
    records: Vec<&'a [u8]>,
}

fn is_named_tag(tag: &[u8], name: &str) -> Result<bool> {
    if wire::record_type(tag)? != PROG_BINARY_TAG {
        return Ok(false);
    }
    let name = wire::utf16(name);
    for child in wire::children(tag)? {
        let child = child?;
        if wire::record_type(child)? == CSTRING {
            return Ok(wire::payload(child)? == name.as_slice());
        }
    }
    Ok(false)
}

fn unsupported(element: &'static str, reason: &'static str) -> EditError {
    EditError::Refused(Refusal::UnsupportedElement { element, reason })
}

/// Locate the tag named `name` in the container `owner` of type
/// `owner_type`. An edit of `element` refuses an ambiguous owner.
pub(crate) fn locate<'a>(
    owner: &'a [u8],
    owner_type: u16,
    name: &'static str,
    element: &'static str,
) -> Result<Blob<'a>> {
    if wire::record_type(owner)? != owner_type {
        return Err(Error::Corrupted(
            "programmable-tag owner has an unexpected record type".into(),
        )
        .into());
    }
    let mut blob = Blob {
        name,
        owner_children: wire::child_list(owner)?,
        prog_tags: None,
        tags: Vec::new(),
        tag: None,
        tag_children: Vec::new(),
        blob: None,
        records: Vec::new(),
    };
    for (index, child) in blob.owner_children.iter().enumerate() {
        if wire::record_type(child)? == PROG_TAGS && blob.prog_tags.replace(index).is_some() {
            return Err(unsupported(
                element,
                "the owner has several programmable-tag containers",
            ));
        }
    }
    let Some(prog_tags) = blob.prog_tags else {
        return Ok(blob);
    };
    blob.tags = wire::child_list(blob.owner_children[prog_tags])?;
    for (index, tag) in blob.tags.iter().enumerate() {
        if is_named_tag(tag, name)? && blob.tag.replace(index).is_some() {
            return Err(unsupported(
                element,
                "the owner has several tags of one version",
            ));
        }
    }
    let Some(tag) = blob.tag else {
        return Ok(blob);
    };
    blob.tag_children = wire::child_list(blob.tags[tag])?;
    for (index, child) in blob.tag_children.iter().enumerate() {
        if wire::record_type(child)? == BINARY_TAG_DATA && blob.blob.replace(index).is_some() {
            return Err(
                Error::Corrupted("a programmable tag has several data blobs".into()).into(),
            );
        }
    }
    let data = blob.blob.ok_or_else(|| {
        EditError::from(Error::Corrupted(
            "a programmable tag has no data blob".into(),
        ))
    })?;
    blob.records = wire::sequence(wire::payload(blob.tag_children[data])?)?;
    Ok(blob)
}

impl<'a> Blob<'a> {
    /// The blob's records in order; empty when the tag is absent.
    pub(crate) fn records(&self) -> &[&'a [u8]] {
        &self.records
    }

    /// Re-frame `owner` around new blob `records`. A missing tag container
    /// goes after the last owner child whose type is in `before_tags`.
    pub(crate) fn rebuild(
        &self,
        owner: &[u8],
        records: &[Vec<u8>],
        before_tags: &[u16],
    ) -> Result<Vec<u8>> {
        let blob_payload = records.concat();
        let tag = match (self.tag, self.blob) {
            (Some(tag), Some(blob)) => {
                let mut children = Vec::with_capacity(self.tag_children.len());
                for (index, child) in self.tag_children.iter().enumerate() {
                    children.push(if index == blob {
                        wire::rebuild(child, &blob_payload)?
                    } else {
                        child.to_vec()
                    });
                }
                Some((tag, wire::rebuild(self.tags[tag], &children.concat())?))
            },
            _ => None,
        };
        let new_tag = || -> Result<Vec<u8>> {
            Ok(wire::container(
                PROG_BINARY_TAG,
                0,
                &[
                    wire::encode(CSTRING, 0, 0, &wire::utf16(self.name))?,
                    wire::encode(BINARY_TAG_DATA, 0, 0, &blob_payload)?,
                ],
            )?)
        };
        let (position, container, inserted) = match self.prog_tags {
            Some(index) => {
                let mut tags = self.tags.iter().map(|tag| tag.to_vec()).collect::<Vec<_>>();
                match tag {
                    Some((position, rewritten)) => tags[position] = rewritten,
                    None => tags.push(new_tag()?),
                }
                (
                    index,
                    wire::rebuild(self.owner_children[index], &tags.concat())?,
                    false,
                )
            },
            None => {
                let mut position = 0;
                for (index, child) in self.owner_children.iter().enumerate() {
                    if before_tags.contains(&wire::record_type(child)?) {
                        position = index + 1;
                    }
                }
                (
                    position,
                    wire::container(PROG_TAGS, 0, &[new_tag()?])?,
                    true,
                )
            },
        };
        let mut data = Vec::with_capacity(owner.len().saturating_add(container.len()));
        for (index, child) in self.owner_children.iter().enumerate() {
            if index == position {
                data.extend_from_slice(&container);
                if !inserted {
                    continue;
                }
            }
            data.extend_from_slice(child);
        }
        if inserted && position == self.owner_children.len() {
            data.extend_from_slice(&container);
        }
        Ok(wire::rebuild(owner, &data)?)
    }
}
