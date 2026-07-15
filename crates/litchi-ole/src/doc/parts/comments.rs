//! Comment-table parsing for Word 97+ binary documents.
//!
//! Comments use references in the main document (`PlcfandRef`), ranges in the
//! comment subdocument (`PlcfandTxt`), and a separate array of author names.

use super::super::package::{DocError, Result};
use super::fib::FileInformationBlock;
use crate::plcf::PlcfParser;

const ATRD_PRE10_SIZE: usize = 30;

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

            let descriptor = CommentDescriptor::parse(
                ref_plcf.property(index).ok_or_else(|| {
                    DocError::Corrupted("PlcfandRef is missing an ATRDPre10".to_string())
                })?,
                owners.len(),
            )?;
            let marker_cp = subdoc_start.checked_add(text_cps[index]).ok_or_else(|| {
                DocError::Corrupted("comment story start CP overflows".to_string())
            })?;
            let text_end_cp = subdoc_start
                .checked_add(text_cps[index + 1])
                .ok_or_else(|| DocError::Corrupted("comment story end CP overflows".to_string()))?;
            let author = owners[usize::from(descriptor.author_index)].clone();

            references.push(CommentReference {
                reference_cp,
                marker_cp,
                text_end_cp,
                author,
                descriptor,
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
        let mut fib_data = vec![0; 512];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[76..80].copy_from_slice(&10u32.to_le_bytes()); // ccpText
        fib_data[92..96].copy_from_slice(&5u32.to_le_bytes()); // ccpAtn

        let mut table = Vec::new();
        let owners_offset = table.len() as u32;
        table.extend_from_slice(&[5, 0, b'A', 0, b'l', 0, b'i', 0, b'c', 0, b'e', 0]);
        set_fib_pointer(&mut fib_data, 36, owners_offset, 12);

        let ref_offset = table.len() as u32;
        table.extend_from_slice(&2u32.to_le_bytes());
        table.extend_from_slice(&10u32.to_le_bytes());
        table.extend_from_slice(&descriptor("AE", 0, 7));
        set_fib_pointer(&mut fib_data, 4, ref_offset, 38);

        let txt_offset = table.len() as u32;
        for cp in [0u32, 4, 99] {
            table.extend_from_slice(&cp.to_le_bytes());
        }
        set_fib_pointer(&mut fib_data, 5, txt_offset, 12);

        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let comments = CommentsTable::parse(&fib, &table).unwrap();
        assert_eq!(comments.count(), 1);
        let reference = comments.find_at_position(2).unwrap();
        assert_eq!(reference.author, "Alice");
        assert_eq!(reference.descriptor.initials, "AE");
        assert_eq!(reference.descriptor.bookmark_tag, Some(7));
        assert_eq!((reference.marker_cp, reference.text_end_cp), (10, 14));
    }
}
