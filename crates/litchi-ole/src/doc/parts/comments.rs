//! Comment-table parsing for Word 97+ binary documents.
//!
//! Comments use references in the main document (`PlcfandRef`), ranges in the
//! comment subdocument (`PlcfandTxt`), and a separate array of author names.

use super::super::package::{DocError, Result};
use super::super::{CommentDateTime, CommentExtendedMetadata};
use super::fib::FileInformationBlock;
use crate::plcf::PlcfParser;
use std::collections::{HashMap, HashSet};

const ATRD_PRE10_SIZE: usize = 30;
const ATRD_POST10_SIZE: usize = 18;

/// The 30-byte descriptor associated with a comment reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentDescriptor {
    /// Initials stored directly in the descriptor.
    pub initials: String,
    /// Index into the comment-owner XST array.
    pub author_index: u16,
    /// Annotation-bookmark tag, or `None` for a zero-length commented range.
    pub bookmark_tag: Option<u32>,
}

impl CommentDescriptor {
    fn parse(data: &[u8], author_count: usize) -> Result<Self> {
        if data.len() != ATRD_PRE10_SIZE {
            return Err(DocError::Corrupted(
                "ATRDPre10 must be exactly 30 bytes".to_string(),
            ));
        }

        let initials_len = usize::from(read_u16(data, 0, "ATRDPre10 initials length")?);
        if initials_len > 9 {
            return Err(DocError::Corrupted(
                "ATRDPre10 initials exceed nine UTF-16 characters".to_string(),
            ));
        }
        let initials = decode_utf16(&data[2..2 + initials_len * 2], "comment initials")?;
        let author_index = read_u16(data, 20, "ATRDPre10 author index")?;
        if usize::from(author_index) >= author_count {
            return Err(DocError::Corrupted(
                "ATRDPre10 author index exceeds the comment-owner array".to_string(),
            ));
        }
        if read_u16(data, 22, "ATRDPre10 unused bits")? != 0
            || read_u16(data, 24, "ATRDPre10 unused flags")? != 0
        {
            return Err(DocError::Corrupted(
                "ATRDPre10 reserved fields must be zero".to_string(),
            ));
        }

        let raw_tag = litchi_core::binary::read_i32_le(data, 26).map_err(|error| {
            DocError::Corrupted(format!("invalid ATRDPre10 bookmark tag: {error}"))
        })?;
        let bookmark_tag = match raw_tag {
            -1 => None,
            value if value >= 0 => Some(value as u32),
            _ => {
                return Err(DocError::Corrupted(
                    "ATRDPre10 bookmark tag is less than -1".to_string(),
                ));
            },
        };

        Ok(Self {
            initials,
            author_index,
            bookmark_tag,
        })
    }
}

/// One comment reference and its range in the comment subdocument.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentReference {
    /// CP of the U+0005 comment-reference character in the main document.
    pub reference_cp: u32,
    /// Absolute CP of the U+0005 marker that begins the comment story.
    pub marker_cp: u32,
    /// Absolute CP immediately after the comment story.
    pub text_end_cp: u32,
    /// Comment author name resolved through the owner array.
    pub author: String,
    /// Descriptor stored in `PlcfandRef`.
    pub descriptor: CommentDescriptor,
    /// Start CP of the commented main-document range, if this is a range comment.
    pub range_start_cp: Option<u32>,
    /// Exclusive end CP of the commented main-document range.
    pub range_end_cp: Option<u32>,
    /// Word 2002+ timestamp, reply-tree, and ink metadata, when present.
    pub extended_metadata: Option<CommentExtendedMetadata>,
}

/// Parsed Word comment tables.
#[derive(Debug, Clone, Default)]
pub struct CommentsTable {
    references: Vec<CommentReference>,
}

impl CommentsTable {
    /// Parse all Word 97+ comment references and text ranges.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let Some((subdoc_start, subdoc_end)) = fib.get_comment_range() else {
            return Ok(Self::default());
        };
        let main_end = fib.get_main_doc_range().1;
        let owners = parse_owners(required_table_slice(
            fib,
            table_stream,
            36,
            "comment-owner array",
        )?)?;
        if owners.is_empty() {
            return Err(DocError::Corrupted(
                "a document with comments must define at least one owner".to_string(),
            ));
        }

        let ref_data = required_table_slice(fib, table_stream, 4, "PlcfandRef")?;
        if ref_data.len() < 4 || (ref_data.len() - 4) % (4 + ATRD_PRE10_SIZE) != 0 {
            return Err(DocError::Corrupted(
                "PlcfandRef has an invalid byte length".to_string(),
            ));
        }
        let ref_plcf = PlcfParser::parse(ref_data, ATRD_PRE10_SIZE)
            .ok_or_else(|| DocError::Corrupted("PlcfandRef is malformed".to_string()))?;
        if ref_plcf.count() == 0 {
            return Err(DocError::Corrupted(
                "a nonempty comment subdocument has no references".to_string(),
            ));
        }

        let txt_data = required_table_slice(fib, table_stream, 5, "PlcfandTxt")?;
        let subdoc_len = subdoc_end.checked_sub(subdoc_start).ok_or_else(|| {
            DocError::Corrupted("comment subdocument range is reversed".to_string())
        })?;
        let text_cps = parse_text_cps(txt_data, ref_plcf.count(), subdoc_len)?;

        let mut descriptors = Vec::with_capacity(ref_plcf.count());
        let mut referenced_tags = HashSet::new();
        for index in 0..ref_plcf.count() {
            let descriptor = CommentDescriptor::parse(
                ref_plcf.property(index).ok_or_else(|| {
                    DocError::Corrupted("PlcfandRef is missing an ATRDPre10".to_string())
                })?,
                owners.len(),
            )?;
            if let Some(tag) = descriptor.bookmark_tag
                && !referenced_tags.insert(tag)
            {
                return Err(DocError::Corrupted(
                    "multiple comments reference the same annotation bookmark tag".to_string(),
                ));
            }
            descriptors.push(descriptor);
        }
        let ranges = parse_annotation_ranges(fib, table_stream, main_end, &referenced_tags)?;
        let extended_metadata = parse_extended_metadata(fib, table_stream, ref_plcf.count())?;

        let mut references = Vec::with_capacity(ref_plcf.count());
        let mut previous_ref = None;
        for index in 0..ref_plcf.count() {
            let reference_cp = ref_plcf.position(index).ok_or_else(|| {
                DocError::Corrupted("PlcfandRef is missing a character position".to_string())
            })?;
            if reference_cp >= main_end || previous_ref.is_some_and(|cp| cp >= reference_cp) {
                return Err(DocError::Corrupted(
                    "PlcfandRef CPs must be unique, increasing, and inside the main document"
                        .to_string(),
                ));
            }
            previous_ref = Some(reference_cp);

            let descriptor = descriptors[index].clone();
            let marker_cp = subdoc_start.checked_add(text_cps[index]).ok_or_else(|| {
                DocError::Corrupted("comment story start CP overflows".to_string())
            })?;
            let text_end_cp = subdoc_start
                .checked_add(text_cps[index + 1])
                .ok_or_else(|| DocError::Corrupted("comment story end CP overflows".to_string()))?;
            let author = owners[usize::from(descriptor.author_index)].clone();
            let range = descriptor
                .bookmark_tag
                .map(|tag| {
                    ranges.get(&tag).copied().ok_or_else(|| {
                        DocError::Corrupted(
                            "ATRDPre10 references an unknown annotation bookmark tag".to_string(),
                        )
                    })
                })
                .transpose()?;

            references.push(CommentReference {
                reference_cp,
                marker_cp,
                text_end_cp,
                author,
                descriptor,
                range_start_cp: range.map(|value| value.0),
                range_end_cp: range.map(|value| value.1),
                extended_metadata: extended_metadata.as_ref().map(|metadata| metadata[index]),
            });
        }

        Ok(Self { references })
    }

    /// All comments in main-document reference order.
    pub fn references(&self) -> &[CommentReference] {
        &self.references
    }

    /// Number of comments.
    pub fn count(&self) -> usize {
        self.references.len()
    }

    /// Find the comment whose reference character is at `cp`.
    pub fn find_at_position(&self, cp: u32) -> Option<&CommentReference> {
        self.references
            .iter()
            .find(|reference| reference.reference_cp == cp)
    }
}

fn parse_extended_metadata(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    comment_count: usize,
) -> Result<Option<Vec<CommentExtendedMetadata>>> {
    let Some((_, length)) = fib.get_table_pointer(112) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let data = required_table_slice(fib, table_stream, 112, "AtrdExtra")?;
    if data.len() != comment_count * ATRD_POST10_SIZE {
        return Err(DocError::Corrupted(
            "AtrdExtra count does not match PlcfandRef".to_string(),
        ));
    }

    let mut metadata = Vec::with_capacity(comment_count);
    let mut parent_deltas = Vec::with_capacity(comment_count);
    for record in data.chunks_exact(ATRD_POST10_SIZE) {
        let packed_time = litchi_core::binary::read_u32_le(record, 0)
            .map_err(|error| DocError::Corrupted(format!("invalid ATRDPost10 DTTM: {error}")))?;
        let modified_at = parse_dttm(packed_time)?;
        if read_u16(record, 4, "ATRDPost10 padding1")? != 0 {
            return Err(DocError::Corrupted(
                "ATRDPost10 padding1 must be zero".to_string(),
            ));
        }
        let depth = litchi_core::binary::read_u32_le(record, 6).map_err(|error| {
            DocError::Corrupted(format!("invalid ATRDPost10 comment depth: {error}"))
        })?;
        let parent_delta = litchi_core::binary::read_i32_le(record, 10).map_err(|error| {
            DocError::Corrupted(format!("invalid ATRDPost10 parent offset: {error}"))
        })?;
        let flags = litchi_core::binary::read_u32_le(record, 14)
            .map_err(|error| DocError::Corrupted(format!("invalid ATRDPost10 flags: {error}")))?;
        if flags & !0x2 != 0 {
            return Err(DocError::Corrupted(
                "ATRDPost10 reserved flag bits must be zero".to_string(),
            ));
        }
        metadata.push(CommentExtendedMetadata {
            modified_at,
            depth,
            parent_index: None,
            is_ink: flags & 0x2 != 0,
        });
        parent_deltas.push(parent_delta);
    }

    let mut active_ancestors = Vec::<usize>::new();
    for index in 0..metadata.len() {
        let depth = usize::try_from(metadata[index].depth).map_err(|_| {
            DocError::Corrupted("ATRDPost10 comment depth is too large".to_string())
        })?;
        if depth > active_ancestors.len() {
            return Err(DocError::Corrupted(
                "AtrdExtra comment depths are not in pre-order".to_string(),
            ));
        }
        active_ancestors.truncate(depth);

        let parent_delta = parent_deltas[index];
        if depth == 0 {
            if parent_delta != 0 {
                return Err(DocError::Corrupted(
                    "a top-level ATRDPost10 must have no parent".to_string(),
                ));
            }
        } else {
            let parent = index
                .checked_add_signed(parent_delta as isize)
                .ok_or_else(|| {
                    DocError::Corrupted("ATRDPost10 parent offset is out of range".to_string())
                })?;
            let expected_parent = active_ancestors.get(depth - 1).copied().ok_or_else(|| {
                DocError::Corrupted("AtrdExtra comment tree is malformed".to_string())
            })?;
            if parent != expected_parent || metadata[parent].depth + 1 != metadata[index].depth {
                return Err(DocError::Corrupted(
                    "ATRDPost10 parent and depth do not describe a pre-order tree".to_string(),
                ));
            }
            metadata[index].parent_index = Some(parent);
        }
        active_ancestors.push(index);
    }

    Ok(Some(metadata))
}

fn parse_dttm(value: u32) -> Result<Option<CommentDateTime>> {
    let minute = (value & 0x3F) as u8;
    let hour = ((value >> 6) & 0x1F) as u8;
    let day = ((value >> 11) & 0x1F) as u8;
    let month = ((value >> 16) & 0x0F) as u8;
    let year = ((value >> 20) & 0x01FF) as u16 + 1900;
    let weekday = ((value >> 29) & 0x07) as u8;
    if minute > 59 || hour > 23 || day > 31 || month > 12 || weekday > 6 {
        return Err(DocError::Corrupted(
            "ATRDPost10 contains an invalid DTTM".to_string(),
        ));
    }
    if day == 0 || month == 0 {
        return Ok(None);
    }
    Ok(Some(CommentDateTime {
        year,
        month,
        day,
        hour,
        minute,
        weekday,
    }))
}

fn parse_annotation_ranges(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    main_end: u32,
    referenced_tags: &HashSet<u32>,
) -> Result<HashMap<u32, (u32, u32)>> {
    let pointer_lengths = [37usize, 42, 43].map(|index| {
        fib.get_table_pointer(index)
            .map(|(_, length)| length)
            .unwrap_or(0)
    });
    if referenced_tags.is_empty() {
        if pointer_lengths.iter().any(|&length| length != 0) {
            return Err(DocError::Corrupted(
                "point comments must not define annotation bookmark tables".to_string(),
            ));
        }
        return Ok(HashMap::new());
    }

    let bookmark_names = required_table_slice(fib, table_stream, 37, "SttbfAtnBkmk")?;
    let tags = parse_annotation_bookmark_tags(bookmark_names)?;
    let starts_data = required_table_slice(fib, table_stream, 42, "PlcfAtnBkf")?;
    if starts_data.len() < 4 || (starts_data.len() - 4) % 8 != 0 {
        return Err(DocError::Corrupted(
            "PlcfAtnBkf has an invalid byte length".to_string(),
        ));
    }
    let starts = PlcfParser::parse(starts_data, 4)
        .ok_or_else(|| DocError::Corrupted("PlcfAtnBkf is malformed".to_string()))?;
    if starts.count() != tags.len() {
        return Err(DocError::Corrupted(
            "PlcfAtnBkf and SttbfAtnBkmk counts do not match".to_string(),
        ));
    }

    validate_bookmark_cps(&starts, main_end, "PlcfAtnBkf")?;

    let ends_data = required_table_slice(fib, table_stream, 43, "PlcfAtnBkl")?;
    if ends_data.len() != (tags.len() + 1) * 4 {
        return Err(DocError::Corrupted(
            "PlcfAtnBkl count does not match PlcfAtnBkf".to_string(),
        ));
    }
    let mut ends = Vec::with_capacity(tags.len() + 1);
    for offset in (0..ends_data.len()).step_by(4) {
        ends.push(
            litchi_core::binary::read_u32_le(ends_data, offset).map_err(|error| {
                DocError::Corrupted(format!("invalid PlcfAtnBkl character position: {error}"))
            })?,
        );
    }
    // The final CP of a bookmark PLC is ignored per [MS-DOC] 2.8.10; writers
    // disagree on whether it counts the paragraph mark that separates the main
    // document from the subdocuments, so it carries no reliable information.
    if ends[..ends.len() - 1].iter().any(|&cp| cp > main_end)
        || ends[..ends.len() - 1]
            .windows(2)
            .any(|pair| pair[0] > pair[1])
    {
        return Err(DocError::Corrupted(
            "PlcfAtnBkl has invalid or non-monotonic character positions".to_string(),
        ));
    }

    let mut used_end_indexes = HashSet::new();
    let mut ranges = HashMap::with_capacity(tags.len());
    for (index, tag) in tags.into_iter().enumerate() {
        let property = starts
            .property(index)
            .ok_or_else(|| DocError::Corrupted("PlcfAtnBkf is missing an FBKF".to_string()))?;
        let end_index = usize::from(read_u16(property, 0, "annotation bookmark ibkl")?);
        let bookmark_flags = read_u16(property, 2, "annotation bookmark BKC")?;
        if bookmark_flags & 0x8080 != 0 {
            return Err(DocError::Corrupted(
                "annotation bookmark BKC has fPub or fCol set".to_string(),
            ));
        }
        if end_index >= ends.len() - 1 || !used_end_indexes.insert(end_index) {
            return Err(DocError::Corrupted(
                "annotation bookmark ibkl values must be unique and in range".to_string(),
            ));
        }
        let start = starts
            .position(index)
            .ok_or_else(|| DocError::Corrupted("PlcfAtnBkf is missing a start CP".to_string()))?;
        let end = ends[end_index];
        if start > end {
            return Err(DocError::Corrupted(
                "annotation bookmark start CP exceeds its end CP".to_string(),
            ));
        }
        ranges.insert(tag, (start, end));
    }
    if ranges.len() != referenced_tags.len()
        || referenced_tags.iter().any(|tag| !ranges.contains_key(tag))
    {
        return Err(DocError::Corrupted(
            "annotation bookmark tags do not match ranged comments".to_string(),
        ));
    }
    Ok(ranges)
}

fn parse_annotation_bookmark_tags(data: &[u8]) -> Result<Vec<u32>> {
    if data.len() < 6 {
        return Err(DocError::Corrupted(
            "SttbfAtnBkmk is missing its header".to_string(),
        ));
    }
    let count = usize::from(read_u16(data, 2, "SttbfAtnBkmk count")?);
    if read_u16(data, 0, "SttbfAtnBkmk fExtend")? != 0xFFFF
        || count > 0x3FFC
        || read_u16(data, 4, "SttbfAtnBkmk cbExtra")? != 10
        || data.len() != 6 + count * 12
    {
        return Err(DocError::Corrupted(
            "SttbfAtnBkmk has invalid header fields or byte length".to_string(),
        ));
    }
    let mut tags = Vec::with_capacity(count);
    let mut unique = HashSet::with_capacity(count);
    for index in 0..count {
        let offset = 6 + index * 12;
        if read_u16(data, offset, "SttbfAtnBkmk string length")? != 0
            || read_u16(data, offset + 2, "ATNBE bookmark class")? != 0x0100
            || litchi_core::binary::read_i32_le(data, offset + 8).map_err(|error| {
                DocError::Corrupted(format!("invalid ATNBE legacy tag: {error}"))
            })? != -1
        {
            return Err(DocError::Corrupted(
                "SttbfAtnBkmk contains an invalid ATNBE".to_string(),
            ));
        }
        let tag = litchi_core::binary::read_u32_le(data, offset + 4)
            .map_err(|error| DocError::Corrupted(format!("invalid ATNBE tag: {error}")))?;
        if !unique.insert(tag) {
            return Err(DocError::Corrupted("ATNBE tags must be unique".to_string()));
        }
        tags.push(tag);
    }
    Ok(tags)
}

fn validate_bookmark_cps(plcf: &PlcfParser, main_end: u32, name: &str) -> Result<()> {
    // Every CP except the last must be inside the main document and
    // monotonic. The final CP of a bookmark PLC is ignored per [MS-DOC]
    // 2.8.10, so no constraint is placed on it.
    let mut previous = None;
    for index in 0..plcf.count() {
        let cp = plcf
            .position(index)
            .ok_or_else(|| DocError::Corrupted(format!("{name} is missing a CP")))?;
        if cp > main_end || previous.is_some_and(|value| value > cp) {
            return Err(DocError::Corrupted(format!(
                "{name} has out-of-range or non-monotonic CPs"
            )));
        }
        previous = Some(cp);
    }
    Ok(())
}

fn required_table_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<&'a [u8]> {
    let (offset, length) = fib
        .get_table_pointer(index)
        .filter(|(_, length)| *length != 0)
        .ok_or_else(|| DocError::Corrupted(format!("{name} is missing")))?;
    let start = usize::try_from(offset)
        .map_err(|_| DocError::Corrupted(format!("{name} offset is too large")))?;
    let length = usize::try_from(length)
        .map_err(|_| DocError::Corrupted(format!("{name} length is too large")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| DocError::Corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .ok_or_else(|| DocError::Corrupted(format!("{name} extends beyond the table stream")))
}

fn parse_owners(data: &[u8]) -> Result<Vec<String>> {
    let mut owners = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let length = usize::from(read_u16(data, offset, "comment-owner XST length")?);
        if length >= 56 {
            return Err(DocError::Corrupted(
                "comment-owner name contains 56 or more UTF-16 characters".to_string(),
            ));
        }
        let byte_length = length
            .checked_mul(2)
            .ok_or_else(|| DocError::Corrupted("comment-owner XST overflows".to_string()))?;
        let start = offset + 2;
        let end = start
            .checked_add(byte_length)
            .ok_or_else(|| DocError::Corrupted("comment-owner XST overflows".to_string()))?;
        let owner = decode_utf16(
            data.get(start..end)
                .ok_or_else(|| DocError::Corrupted("comment-owner XST is truncated".to_string()))?,
            "comment-owner name",
        )?;
        if owners.iter().any(|existing| existing == &owner) {
            return Err(DocError::Corrupted(
                "comment-owner names must be unique".to_string(),
            ));
        }
        owners.push(owner);
        offset = end;
    }
    if owners.len() > 0x7FFF {
        return Err(DocError::Corrupted(
            "comment-owner array exceeds 0x7FFF entries".to_string(),
        ));
    }
    Ok(owners)
}

fn parse_text_cps(data: &[u8], comment_count: usize, subdoc_len: u32) -> Result<Vec<u32>> {
    if data.len() % 4 != 0 || data.len() / 4 != comment_count + 2 {
        return Err(DocError::Corrupted(
            "PlcfandTxt CP count does not match PlcfandRef".to_string(),
        ));
    }
    let mut cps = Vec::with_capacity(comment_count + 2);
    for offset in (0..data.len()).step_by(4) {
        cps.push(
            litchi_core::binary::read_u32_le(data, offset).map_err(|error| {
                DocError::Corrupted(format!("invalid PlcfandTxt character position: {error}"))
            })?,
        );
    }
    if subdoc_len == 0
        || cps[..cps.len() - 1].iter().any(|&cp| cp >= subdoc_len)
        || cps[..cps.len() - 1]
            .windows(2)
            .any(|pair| pair[0] >= pair[1])
        || cps[..cps.len() - 1]
            .windows(2)
            .any(|pair| pair[1] - pair[0] < 2)
    {
        return Err(DocError::Corrupted(
            "PlcfandTxt CPs must be unique, increasing, and inside the comment subdocument"
                .to_string(),
        ));
    }
    if cps[cps.len() - 2] != subdoc_len - 1 {
        return Err(DocError::Corrupted(
            "PlcfandTxt terminator must equal ccpAtn - 1".to_string(),
        ));
    }
    Ok(cps)
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| DocError::Corrupted(format!("invalid {field}: {error}")))
}

fn decode_utf16(data: &[u8], field: &str) -> Result<String> {
    if data.len() % 2 != 0 {
        return Err(DocError::Corrupted(format!(
            "{field} contains a partial UTF-16 code unit"
        )));
    }
    let units = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units)
        .map_err(|_| DocError::Corrupted(format!("{field} contains invalid UTF-16")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn set_fib_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
        let declared = u16::from_le_bytes([fib[152], fib[153]]);
        let count = declared.max(u16::try_from(index + 1).unwrap());
        fib[152..154].copy_from_slice(&count.to_le_bytes());
        let start = 154 + index * 8;
        fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
        fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
    }

    fn descriptor(initials: &str, author_index: u16, tag: i32) -> Vec<u8> {
        let units = initials.encode_utf16().collect::<Vec<_>>();
        let mut data = vec![0; ATRD_PRE10_SIZE];
        data[0..2].copy_from_slice(&(units.len() as u16).to_le_bytes());
        for (index, unit) in units.into_iter().enumerate() {
            data[2 + index * 2..4 + index * 2].copy_from_slice(&unit.to_le_bytes());
        }
        data[20..22].copy_from_slice(&author_index.to_le_bytes());
        data[26..30].copy_from_slice(&tag.to_le_bytes());
        data
    }

    fn dttm(year: u16, month: u8, day: u8, hour: u8, minute: u8, weekday: u8) -> u32 {
        u32::from(minute)
            | (u32::from(hour) << 6)
            | (u32::from(day) << 11)
            | (u32::from(month) << 16)
            | (u32::from(year - 1900) << 20)
            | (u32::from(weekday) << 29)
    }

    fn extended_record(timestamp: u32, depth: u32, parent_delta: i32, flags: u32) -> Vec<u8> {
        let mut data = Vec::with_capacity(ATRD_POST10_SIZE);
        data.extend_from_slice(&timestamp.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&depth.to_le_bytes());
        data.extend_from_slice(&parent_delta.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data
    }

    #[test]
    fn parses_unicode_owners_and_descriptors() {
        let owners = [3, 0, b'A', 0, 0x3D, 0xD8, 2, 0xDE]; // A + U+1F602
        assert_eq!(parse_owners(&owners).unwrap(), ["A😂"]);

        let parsed = CommentDescriptor::parse(&descriptor("XY", 0, -1), 1).unwrap();
        assert_eq!(parsed.initials, "XY");
        assert_eq!(parsed.author_index, 0);
        assert_eq!(parsed.bookmark_tag, None);
    }

    #[test]
    fn rejects_invalid_owner_arrays_and_descriptors() {
        assert!(parse_owners(&[1, 0, 0x00]).is_err());
        assert!(parse_owners(&[1, 0, b'A', 0, 1, 0, b'A', 0]).is_err());
        assert!(CommentDescriptor::parse(&descriptor("X", 1, -1), 1).is_err());
        assert!(CommentDescriptor::parse(&descriptor("X", 0, -2), 1).is_err());

        let mut reserved = descriptor("X", 0, -1);
        reserved[22] = 1;
        assert!(CommentDescriptor::parse(&reserved, 1).is_err());
    }

    #[test]
    fn validates_comment_text_character_positions() {
        let encode = |cps: &[u32]| {
            cps.iter()
                .flat_map(|cp| cp.to_le_bytes())
                .collect::<Vec<_>>()
        };
        assert_eq!(
            parse_text_cps(&encode(&[0, 6, 7]), 1, 7).unwrap(),
            [0, 6, 7]
        );
        assert!(parse_text_cps(&encode(&[0, 0, 6]), 1, 7).is_err());
        assert!(parse_text_cps(&encode(&[0, 1, 2]), 1, 2).is_err());
        assert!(parse_text_cps(&encode(&[0, 4, 6]), 1, 8).is_err());
        assert!(parse_text_cps(&[0; 11], 1, 7).is_err());
    }

    #[test]
    fn parses_complete_comment_tables_from_fib_pointers() {
        let mut fib_data = vec![0; 154 + 136 * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        fib_data[76..80].copy_from_slice(&10u32.to_le_bytes()); // ccpText
        fib_data[92..96].copy_from_slice(&5u32.to_le_bytes()); // ccpAtn
        fib_data[152..154].copy_from_slice(&136u16.to_le_bytes());

        let mut table = Vec::new();
        let owners_offset = table.len() as u32;
        table.extend_from_slice(&[5, 0, b'A', 0, b'l', 0, b'i', 0, b'c', 0, b'e', 0]);
        set_fib_pointer(&mut fib_data, 36, owners_offset, 12);

        let ref_offset = table.len() as u32;
        table.extend_from_slice(&2u32.to_le_bytes());
        table.extend_from_slice(&10u32.to_le_bytes());
        table.extend_from_slice(&descriptor("AE", 0, -1));
        set_fib_pointer(&mut fib_data, 4, ref_offset, 38);

        let txt_offset = table.len() as u32;
        for cp in [0u32, 4, 99] {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        set_fib_pointer(&mut fib_data, 5, txt_offset, 12);

        let metadata_offset = table.len() as u32;
        let timestamp = dttm(2026, 7, 15, 10, 30, 3);
        table.extend_from_slice(&extended_record(timestamp, 0, 0, 0x2));
        set_fib_pointer(&mut fib_data, 112, metadata_offset, ATRD_POST10_SIZE as u32);

        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let comments = CommentsTable::parse(&fib, &table).unwrap();
        assert_eq!(comments.count(), 1);
        let reference = comments.find_at_position(2).unwrap();
        assert_eq!(reference.author, "Alice");
        assert_eq!(reference.descriptor.initials, "AE");
        assert_eq!(reference.descriptor.bookmark_tag, None);
        assert_eq!(reference.range_start_cp, None);
        assert_eq!(reference.range_end_cp, None);
        assert_eq!(
            reference.extended_metadata,
            Some(CommentExtendedMetadata {
                modified_at: Some(CommentDateTime {
                    year: 2026,
                    month: 7,
                    day: 15,
                    hour: 10,
                    minute: 30,
                    weekday: 3,
                }),
                depth: 0,
                parent_index: None,
                is_ink: true,
            })
        );
        assert_eq!((reference.marker_cp, reference.text_end_cp), (10, 14));
    }

    #[test]
    fn resolves_annotation_bookmark_tags_to_main_document_ranges() {
        let mut fib_data = vec![0; 1100];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[76..80].copy_from_slice(&20u32.to_le_bytes());
        fib_data[92..96].copy_from_slice(&5u32.to_le_bytes());
        fib_data[152..154].copy_from_slice(&93u16.to_le_bytes());

        let mut table = Vec::new();
        let owners_offset = table.len() as u32;
        table.extend_from_slice(&[5, 0, b'A', 0, b'l', 0, b'i', 0, b'c', 0, b'e', 0]);
        set_fib_pointer(&mut fib_data, 36, owners_offset, 12);

        let ref_offset = table.len() as u32;
        table.extend_from_slice(&2u32.to_le_bytes());
        table.extend_from_slice(&20u32.to_le_bytes());
        table.extend_from_slice(&descriptor("AE", 0, 7));
        set_fib_pointer(&mut fib_data, 4, ref_offset, 38);

        let txt_offset = table.len() as u32;
        for cp in [0u32, 4, 99] {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        set_fib_pointer(&mut fib_data, 5, txt_offset, 12);

        let names_offset = table.len() as u32;
        table.extend_from_slice(&0xFFFFu16.to_le_bytes());
        table.extend_from_slice(&1u16.to_le_bytes());
        table.extend_from_slice(&10u16.to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes());
        table.extend_from_slice(&0x0100u16.to_le_bytes());
        table.extend_from_slice(&7u32.to_le_bytes());
        table.extend_from_slice(&(-1i32).to_le_bytes());
        set_fib_pointer(&mut fib_data, 37, names_offset, 18);

        let starts_offset = table.len() as u32;
        table.extend_from_slice(&3u32.to_le_bytes());
        table.extend_from_slice(&21u32.to_le_bytes());
        table.extend_from_slice(&0u16.to_le_bytes()); // ibkl
        table.extend_from_slice(&0u16.to_le_bytes()); // BKC
        set_fib_pointer(&mut fib_data, 42, starts_offset, 12);

        let ends_offset = table.len() as u32;
        table.extend_from_slice(&9u32.to_le_bytes());
        table.extend_from_slice(&21u32.to_le_bytes());
        set_fib_pointer(&mut fib_data, 43, ends_offset, 8);

        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let tags = HashSet::from([7]);
        let ranges = parse_annotation_ranges(&fib, &table, 20, &tags).unwrap();
        assert_eq!(ranges.get(&7), Some(&(3, 9)));
        let comments = CommentsTable::parse(&fib, &table).unwrap();
        let comment = comments.find_at_position(2).unwrap();
        assert_eq!(
            (comment.range_start_cp, comment.range_end_cp),
            (Some(3), Some(9))
        );

        // Word 97 and newer writers place the ignored final CP one past the
        // paragraph mark that follows the main document (main_end + 2); it
        // must not be validated ([MS-DOC] 2.8.10).
        let mut word_final_cp = table.clone();
        word_final_cp[starts_offset as usize + 4..starts_offset as usize + 8]
            .copy_from_slice(&22u32.to_le_bytes());
        word_final_cp[ends_offset as usize + 4..ends_offset as usize + 8]
            .copy_from_slice(&22u32.to_le_bytes());
        assert_eq!(
            parse_annotation_ranges(&fib, &word_final_cp, 20, &tags)
                .unwrap()
                .get(&7),
            Some(&(3, 9))
        );

        let mut malformed = table.clone();
        malformed[starts_offset as usize..starts_offset as usize + 4]
            .copy_from_slice(&21u32.to_le_bytes());
        assert!(parse_annotation_ranges(&fib, &malformed, 20, &tags).is_err());

        let mut bad_bkc = table.clone();
        bad_bkc[starts_offset as usize + 10] = 0x80;
        assert!(parse_annotation_ranges(&fib, &bad_bkc, 20, &tags).is_err());
    }

    #[test]
    fn parses_and_validates_extended_comment_metadata() {
        let mut fib_data = vec![0; 154 + 136 * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        fib_data[152..154].copy_from_slice(&136u16.to_le_bytes());
        let mut table = Vec::new();
        let timestamp = dttm(2026, 7, 15, 10, 30, 3);
        table.extend_from_slice(&extended_record(timestamp, 0, 0, 0));
        table.extend_from_slice(&extended_record(0, 1, -1, 0x2));
        table.extend_from_slice(&extended_record(timestamp, 0, 0, 0));
        set_fib_pointer(&mut fib_data, 112, 0, table.len() as u32);

        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let metadata = parse_extended_metadata(&fib, &table, 3).unwrap().unwrap();
        assert_eq!(
            metadata[0].modified_at,
            Some(CommentDateTime {
                year: 2026,
                month: 7,
                day: 15,
                hour: 10,
                minute: 30,
                weekday: 3,
            })
        );
        assert_eq!(metadata[0].parent_index, None);
        assert_eq!(metadata[1].modified_at, None);
        assert_eq!(metadata[1].depth, 1);
        assert_eq!(metadata[1].parent_index, Some(0));
        assert!(metadata[1].is_ink);

        assert!(parse_extended_metadata(&fib, &table, 2).is_err());

        let mut bad_flags = table.clone();
        bad_flags[14] = 1;
        assert!(parse_extended_metadata(&fib, &bad_flags, 3).is_err());

        let mut bad_parent = table.clone();
        bad_parent[ATRD_POST10_SIZE + 10..ATRD_POST10_SIZE + 14]
            .copy_from_slice(&0i32.to_le_bytes());
        assert!(parse_extended_metadata(&fib, &bad_parent, 3).is_err());

        let mut bad_time = table.clone();
        bad_time[0..4].copy_from_slice(&63u32.to_le_bytes());
        assert!(parse_extended_metadata(&fib, &bad_time, 3).is_err());
    }
}
