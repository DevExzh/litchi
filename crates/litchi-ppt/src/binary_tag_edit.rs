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

    /// Return the exact owner size after appending `extra_record_bytes` to
    /// this blob, without allocating any of the replacement record vectors.
    ///
    /// Hyperlink screen-tip edits use this admission before cloning the
    /// existing binary-tag records.  Every framing layer is fixed-width, so
    /// the projected size can be calculated from the retained source spans.
    pub(crate) fn projected_rebuild_len(
        &self,
        owner: &[u8],
        extra_record_bytes: usize,
    ) -> Result<usize> {
        let records_len = self.records.iter().try_fold(0usize, |total, record| {
            total
                .checked_add(record.len())
                .ok_or_else(|| Error::ResourceLimit("binary-tag record length overflows".into()))
        })?;
        let blob_payload_len = records_len
            .checked_add(extra_record_bytes)
            .ok_or_else(|| Error::ResourceLimit("binary-tag payload length overflows".into()))?;
        let new_blob_len = blob_payload_len
            .checked_add(8)
            .ok_or_else(|| Error::ResourceLimit("BinaryTagData length overflows".into()))?;
        let new_tag_len = match (self.tag, self.blob) {
            (Some(tag), Some(blob)) => self.tags[tag]
                .len()
                .checked_sub(self.tag_children[blob].len())
                .and_then(|payload_without_blob| payload_without_blob.checked_add(new_blob_len))
                .ok_or_else(|| Error::ResourceLimit("ProgBinaryTag length overflows".into()))?,
            (Some(_), None) => {
                return Err(Error::Corrupted(
                    "a programmable tag has no BinaryTagData child".into(),
                )
                .into());
            },
            (None, _) => {
                let name_bytes = self
                    .name
                    .encode_utf16()
                    .count()
                    .checked_mul(2)
                    .ok_or_else(|| Error::ResourceLimit("binary-tag name overflows".into()))?;
                8usize
                    .checked_add(8)
                    .and_then(|total| total.checked_add(name_bytes))
                    .and_then(|total| total.checked_add(new_blob_len))
                    .ok_or_else(|| Error::ResourceLimit("ProgBinaryTag length overflows".into()))?
            },
        };
        let new_prog_tags_len = match (self.prog_tags, self.tag) {
            (Some(prog_tags), Some(tag)) => self.owner_children[prog_tags]
                .len()
                .checked_sub(self.tags[tag].len())
                .and_then(|payload_without_tag| payload_without_tag.checked_add(new_tag_len))
                .ok_or_else(|| Error::ResourceLimit("ProgTags length overflows".into()))?,
            (Some(prog_tags), None) => self.owner_children[prog_tags]
                .len()
                .checked_add(new_tag_len)
                .ok_or_else(|| Error::ResourceLimit("ProgTags length overflows".into()))?,
            (None, _) => 8usize
                .checked_add(new_tag_len)
                .ok_or_else(|| Error::ResourceLimit("ProgTags length overflows".into()))?,
        };
        match self.prog_tags {
            Some(prog_tags) => owner
                .len()
                .checked_sub(self.owner_children[prog_tags].len())
                .and_then(|remaining| remaining.checked_add(new_prog_tags_len))
                .ok_or_else(|| Error::ResourceLimit("owner length overflows".into()))
                .map_err(Into::into),
            None => owner
                .len()
                .checked_add(new_prog_tags_len)
                .ok_or_else(|| Error::ResourceLimit("owner length overflows".into()))
                .map_err(Into::into),
        }
    }

    /// Return the aggregate record-payload charge after appending one
    /// serialized record to this blob.  The charge includes the payload of
    /// every newly materialized framing container as well as the inserted
    /// record's own nested payloads.
    pub(crate) fn projected_rebuild_aggregate(
        &self,
        owner: &[u8],
        extra_record_bytes: usize,
        extra_aggregate: usize,
    ) -> Result<usize> {
        let records_len = self.records.iter().try_fold(0usize, |total, record| {
            total
                .checked_add(record.len())
                .ok_or_else(|| Error::ResourceLimit("binary-tag record length overflows".into()))
        })?;
        let records_aggregate = aggregate_records(&self.records)?;
        let blob_payload_len = records_len
            .checked_add(extra_record_bytes)
            .ok_or_else(|| Error::ResourceLimit("binary-tag payload length overflows".into()))?;
        let blob_aggregate = blob_payload_len
            .checked_add(records_aggregate)
            .and_then(|value| value.checked_add(extra_aggregate))
            .ok_or_else(|| Error::ResourceLimit("BinaryTagData aggregate overflows".into()))?;
        let new_blob_len = blob_payload_len
            .checked_add(8)
            .ok_or_else(|| Error::ResourceLimit("BinaryTagData length overflows".into()))?;
        let new_blob_aggregate = blob_aggregate;

        let (new_tag_len, new_tag_aggregate) = match (self.tag, self.blob) {
            (Some(tag), Some(blob)) => {
                let payload_len = self.tags[tag]
                    .len()
                    .checked_sub(8)
                    .and_then(|payload| payload.checked_sub(self.tag_children[blob].len()))
                    .and_then(|payload_without_blob| payload_without_blob.checked_add(new_blob_len))
                    .ok_or_else(|| Error::ResourceLimit("ProgBinaryTag length overflows".into()))?;
                let other_aggregate = aggregate_indexed(&self.tag_children, blob)?;
                let aggregate = payload_len
                    .checked_add(other_aggregate)
                    .and_then(|value| value.checked_add(new_blob_aggregate))
                    .ok_or_else(|| {
                        Error::ResourceLimit("ProgBinaryTag aggregate overflows".into())
                    })?;
                (
                    payload_len.checked_add(8).ok_or_else(|| {
                        Error::ResourceLimit("ProgBinaryTag length overflows".into())
                    })?,
                    aggregate,
                )
            },
            (Some(_), None) => {
                return Err(Error::Corrupted(
                    "a programmable tag has no BinaryTagData child".into(),
                )
                .into());
            },
            (None, _) => {
                let name_bytes = self
                    .name
                    .encode_utf16()
                    .count()
                    .checked_mul(2)
                    .ok_or_else(|| Error::ResourceLimit("binary-tag name overflows".into()))?;
                let payload_len = 16usize
                    .checked_add(name_bytes)
                    .and_then(|value| value.checked_add(blob_payload_len))
                    .ok_or_else(|| Error::ResourceLimit("ProgBinaryTag length overflows".into()))?;
                let aggregate = payload_len
                    .checked_add(name_bytes)
                    .and_then(|value| value.checked_add(new_blob_aggregate))
                    .ok_or_else(|| {
                        Error::ResourceLimit("ProgBinaryTag aggregate overflows".into())
                    })?;
                (
                    payload_len.checked_add(8).ok_or_else(|| {
                        Error::ResourceLimit("ProgBinaryTag length overflows".into())
                    })?,
                    aggregate,
                )
            },
        };

        let (new_prog_len, new_prog_aggregate) = match (self.prog_tags, self.tag) {
            (Some(prog_tags), Some(tag)) => {
                let payload_len = self.owner_children[prog_tags]
                    .len()
                    .checked_sub(8)
                    .and_then(|payload| payload.checked_sub(self.tags[tag].len()))
                    .and_then(|payload_without_tag| payload_without_tag.checked_add(new_tag_len))
                    .ok_or_else(|| Error::ResourceLimit("ProgTags length overflows".into()))?;
                let other_aggregate = aggregate_indexed(&self.tags, tag)?;
                let aggregate = payload_len
                    .checked_add(other_aggregate)
                    .and_then(|value| value.checked_add(new_tag_aggregate))
                    .ok_or_else(|| Error::ResourceLimit("ProgTags aggregate overflows".into()))?;
                (
                    payload_len
                        .checked_add(8)
                        .ok_or_else(|| Error::ResourceLimit("ProgTags length overflows".into()))?,
                    aggregate,
                )
            },
            (Some(prog_tags), None) => {
                let payload_len = self.owner_children[prog_tags]
                    .len()
                    .checked_sub(8)
                    .and_then(|payload| payload.checked_add(new_tag_len))
                    .ok_or_else(|| Error::ResourceLimit("ProgTags length overflows".into()))?;
                let old_tags_aggregate = aggregate_records(&self.tags)?;
                let aggregate = payload_len
                    .checked_add(old_tags_aggregate)
                    .and_then(|value| value.checked_add(new_tag_aggregate))
                    .ok_or_else(|| Error::ResourceLimit("ProgTags aggregate overflows".into()))?;
                (
                    payload_len
                        .checked_add(8)
                        .ok_or_else(|| Error::ResourceLimit("ProgTags length overflows".into()))?,
                    aggregate,
                )
            },
            (None, _) => {
                let aggregate = new_tag_len
                    .checked_add(new_tag_aggregate)
                    .ok_or_else(|| Error::ResourceLimit("ProgTags aggregate overflows".into()))?;
                (
                    new_tag_len
                        .checked_add(8)
                        .ok_or_else(|| Error::ResourceLimit("ProgTags length overflows".into()))?,
                    aggregate,
                )
            },
        };
        let owner_children_aggregate = match self.prog_tags {
            Some(prog_tags) => aggregate_indexed(&self.owner_children, prog_tags)?,
            None => aggregate_records(&self.owner_children)?,
        };
        let owner_payload_len = match self.prog_tags {
            Some(prog_tags) => owner
                .len()
                .checked_sub(8)
                .and_then(|payload| payload.checked_sub(self.owner_children[prog_tags].len()))
                .and_then(|remaining| remaining.checked_add(new_prog_len))
                .ok_or_else(|| Error::ResourceLimit("owner length overflows".into()))?,
            None => owner
                .len()
                .checked_sub(8)
                .and_then(|payload| payload.checked_add(new_prog_len))
                .ok_or_else(|| Error::ResourceLimit("owner length overflows".into()))?,
        };
        Ok(owner_payload_len
            .checked_add(owner_children_aggregate)
            .and_then(|value| value.checked_add(new_prog_aggregate))
            .ok_or_else(|| Error::ResourceLimit("owner aggregate overflows".into()))?)
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

fn aggregate_records(records: &[&[u8]]) -> Result<usize> {
    records.iter().try_fold(0usize, |total, record| {
        total
            .checked_add(serialized_payload_budget(record)?)
            .ok_or_else(|| {
                EditError::from(Error::ResourceLimit(
                    "serialized-record aggregate overflows".into(),
                ))
            })
    })
}

fn aggregate_indexed(records: &[&[u8]], skip: usize) -> Result<usize> {
    records
        .iter()
        .enumerate()
        .filter(|(index, _)| *index != skip)
        .try_fold(0usize, |total, (_, record)| {
            total
                .checked_add(serialized_payload_budget(record)?)
                .ok_or_else(|| {
                    EditError::from(Error::ResourceLimit(
                        "serialized-record aggregate overflows".into(),
                    ))
                })
        })
}

/// Compute the payload charge used by the strict PPT record parser for one
/// serialized record and all of its nested container records.
pub(crate) fn serialized_payload_budget(record: &[u8]) -> Result<usize> {
    let mut pending = Vec::new();
    pending
        .try_reserve(1)
        .map_err(|_error| Error::AllocationFailed("serialized-record traversal stack"))?;
    pending.push((record, 0usize));
    let mut total = 0usize;
    while let Some((current, depth)) = pending.pop() {
        if depth > crate::SourceLimits::MAX_DEPTH {
            return Err(Error::Limit {
                kind: crate::SourceLimitKind::Depth,
                observed: depth,
                maximum: crate::SourceLimits::MAX_DEPTH,
            }
            .into());
        }
        let payload = wire::payload(current)?;
        total = total
            .checked_add(payload.len())
            .ok_or_else(|| Error::ResourceLimit("serialized-record aggregate overflows".into()))?;
        let record_type = RecordType::from(wire::record_type(current)?);
        if crate::records::Record::is_container_record(record_type) {
            let child_depth = depth
                .checked_add(1)
                .ok_or_else(|| Error::ResourceLimit("serialized-record depth overflows".into()))?;
            for child in wire::children(current)? {
                pending
                    .try_reserve(1)
                    .map_err(|_error| Error::AllocationFailed("serialized-record stack"))?;
                pending.push((child?, child_depth));
            }
        }
    }
    Ok(total)
}
