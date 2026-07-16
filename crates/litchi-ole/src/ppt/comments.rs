//! Strict PowerPoint 10 presentation comment parsing.

use super::package::{PptError, Result};
use super::records::PptRecord;
use super::slide::ParsedComment;
use crate::consts::PptRecordType;

/// Document-level metadata for one presentation-comment author.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointCommentAuthor {
    /// Optional author display name.
    pub name: Option<String>,
    /// Optional zero-based application-defined display color index.
    pub color_index: Option<i32>,
    /// Optional seed for the next comment index created by this author.
    pub comment_index_seed: Option<i32>,
}

/// PowerPoint 10 presentation-comment authors.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointCommentAuthors {
    pub authors: Vec<PowerPointCommentAuthor>,
}

impl PowerPointCommentAuthors {
    /// Parse comment-author records from `___PPT10` document extensions below `root`.
    pub fn parse(root: &PptRecord) -> Result<Self> {
        let mut authors = Vec::new();
        for record in root.versioned_binary_tag_records(10)? {
            if record.record_type == PptRecordType::CommentIndex10 {
                authors.push(parse_author(&record)?);
            }
        }
        Ok(Self { authors })
    }

    /// Find the first author with the specified display name.
    pub fn find(&self, name: &str) -> Option<&PowerPointCommentAuthor> {
        self.authors
            .iter()
            .find(|author| author.name.as_deref() == Some(name))
    }

    /// Validate author index seeds against a collection of parsed slide comments.
    pub fn validate_comments(&self, comments: &[ParsedComment]) -> Result<()> {
        for author in &self.authors {
            let (Some(name), Some(seed)) = (&author.name, author.comment_index_seed) else {
                continue;
            };
            if comments
                .iter()
                .any(|comment| comment.author == *name && comment.index > seed)
            {
                return Err(PptError::Corrupted(format!(
                    "Comment index exceeds the seed for author {name:?}"
                )));
            }
        }
        Ok(())
    }
}

pub(crate) fn parse_slide_comments(root: &PptRecord) -> Result<Vec<ParsedComment>> {
    let mut comments = Vec::new();
    for record in root.versioned_binary_tag_records(10)? {
        if record.record_type == PptRecordType::Comment2000 {
            comments.push(parse_comment(&record)?);
        }
    }
    Ok(comments)
}

fn parse_author(record: &PptRecord) -> Result<PowerPointCommentAuthor> {
    if record.record_type != PptRecordType::CommentIndex10
        || record.version != 0x0f
        || record.instance != 0
    {
        return Err(PptError::Corrupted(
            "CommentIndex10Container has an invalid record header".to_string(),
        ));
    }
    let children = PptRecord::parse_sequence_strict(&record.data, "comment author")?;
    let mut name = None;
    let mut color_index = None;
    let mut comment_index_seed = None;
    for child in children {
        match child.record_type {
            PptRecordType::CString if color_index.is_none() && name.is_none() => {
                name = Some(parse_string(&child, 0, 104, false, "AuthorNameAtom")?);
            },
            PptRecordType::CommentIndex10Atom
                if color_index.is_none() && comment_index_seed.is_none() =>
            {
                if child.version != 0 || child.instance != 0 || child.data.len() != 8 {
                    return Err(PptError::Corrupted(
                        "CommentIndex10Atom has an invalid record header or size".to_string(),
                    ));
                }
                let color = i32::from_le_bytes(child.data[0..4].try_into().map_err(|_| {
                    PptError::Corrupted("Comment color index is truncated".to_string())
                })?);
                let seed = i32::from_le_bytes(child.data[4..8].try_into().map_err(|_| {
                    PptError::Corrupted("Comment index seed is truncated".to_string())
                })?);
                if color < 0 || seed < 0 {
                    return Err(PptError::Corrupted(
                        "Comment author color index or seed is negative".to_string(),
                    ));
                }
                color_index = Some(color);
                comment_index_seed = Some(seed);
            },
            _ => {
                return Err(PptError::Corrupted(
                    "CommentIndex10Container has duplicate, out-of-order, or unexpected children"
                        .to_string(),
                ));
            },
        }
    }
    Ok(PowerPointCommentAuthor {
        name,
        color_index,
        comment_index_seed,
    })
}

fn parse_comment(record: &PptRecord) -> Result<ParsedComment> {
    if record.record_type != PptRecordType::Comment2000
        || record.version != 0x0f
        || record.instance != 0
    {
        return Err(PptError::Corrupted(
            "Comment10Container has an invalid record header".to_string(),
        ));
    }
    let children = PptRecord::parse_sequence_strict(&record.data, "presentation comment")?;
    let mut comment = ParsedComment::default();
    let mut stage = 0u8;
    let mut has_atom = false;
    for child in children {
        match (child.record_type, child.instance) {
            (PptRecordType::CString, 0) if stage == 0 => {
                comment.author = parse_string(&child, 0, 104, false, "Comment10AuthorAtom")?;
                stage = 1;
            },
            (PptRecordType::CString, 1) if stage <= 1 => {
                comment.text = parse_string(&child, 1, 64_000, true, "Comment10TextAtom")?;
                stage = 2;
            },
            (PptRecordType::CString, 2) if stage <= 2 => {
                comment.initials =
                    parse_string(&child, 2, 104, false, "Comment10AuthorInitialAtom")?;
                stage = 3;
            },
            (PptRecordType::Comment2000Atom, _) if !has_atom => {
                parse_comment_atom(&child, &mut comment)?;
                has_atom = true;
                stage = 4;
            },
            _ => {
                return Err(PptError::Corrupted(
                    "Comment10Container has duplicate, out-of-order, or unexpected children"
                        .to_string(),
                ));
            },
        }
    }
    if !has_atom || stage != 4 {
        return Err(PptError::Corrupted(
            "Comment10Container is missing Comment10Atom".to_string(),
        ));
    }
    Ok(comment)
}

fn parse_comment_atom(record: &PptRecord, comment: &mut ParsedComment) -> Result<()> {
    if record.record_type != PptRecordType::Comment2000Atom
        || record.version != 0
        || record.instance != 0
        || record.data.len() != 28
    {
        return Err(PptError::Corrupted(
            "Comment10Atom has an invalid record header or size".to_string(),
        ));
    }
    let data = &record.data;
    comment.index = i32::from_le_bytes(
        data[0..4]
            .try_into()
            .map_err(|_| PptError::Corrupted("Comment10Atom index is truncated".to_string()))?,
    );
    if comment.index < 0 {
        return Err(PptError::Corrupted(
            "Comment10Atom index is negative".to_string(),
        ));
    }
    comment.year = u16::from_le_bytes([data[4], data[5]]);
    comment.month = u16::from_le_bytes([data[6], data[7]]);
    comment.day_of_week = u16::from_le_bytes([data[8], data[9]]);
    comment.day = u16::from_le_bytes([data[10], data[11]]);
    comment.hour = u16::from_le_bytes([data[12], data[13]]);
    comment.minute = u16::from_le_bytes([data[14], data[15]]);
    comment.second = u16::from_le_bytes([data[16], data[17]]);
    comment.millisecond = u16::from_le_bytes([data[18], data[19]]);
    comment.x = i32::from_le_bytes(data[20..24].try_into().map_err(|_| {
        PptError::Corrupted("Comment10Atom horizontal anchor is truncated".to_string())
    })?);
    comment.y = i32::from_le_bytes(data[24..28].try_into().map_err(|_| {
        PptError::Corrupted("Comment10Atom vertical anchor is truncated".to_string())
    })?);
    Ok(())
}

fn parse_string(
    record: &PptRecord,
    instance: u16,
    max_len: usize,
    allow_tab_cr_lf: bool,
    name: &str,
) -> Result<String> {
    if record.record_type != PptRecordType::CString
        || record.version != 0
        || record.instance != instance
        || record.data.len() > max_len
        || record.data.len() & 1 != 0
    {
        return Err(PptError::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }
    let mut units = Vec::with_capacity(record.data.len() / 2);
    for bytes in record.data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        let forbidden = if allow_tab_cr_lf {
            matches!(unit, 0x0001..=0x0008 | 0x000b..=0x000c | 0x000e..=0x001f | 0x007f..=0x009f)
        } else {
            matches!(unit, 0x0001..=0x001f | 0x007f..=0x009f)
        };
        if forbidden {
            return Err(PptError::Corrupted(format!(
                "{name} contains a non-printable character"
            )));
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted(format!("{name} contains invalid UTF-16")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn utf16(value: &str) -> Vec<u8> {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn prog_tags_record(version: u8, blob_payload: &[u8]) -> PptRecord {
        let name = record_bytes(0, 0, 4026, &utf16(&format!("___PPT{version}")));
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        PptRecord {
            record_type: PptRecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: tag.len() as u32,
            data: tag,
            children: Vec::new(),
        }
    }

    fn root(children: Vec<PptRecord>) -> PptRecord {
        PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    fn comment_atom(index: i32) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&index.to_le_bytes());
        for value in [2026, 7, 4, 16, 13, 14, 15, 999] {
            data.extend_from_slice(&u16::try_from(value).unwrap().to_le_bytes());
        }
        data.extend_from_slice(&(-12i32).to_le_bytes());
        data.extend_from_slice(&34i32.to_le_bytes());
        record_bytes(0, 0, 12001, &data)
    }

    fn comment_container(index: i32) -> Vec<u8> {
        let mut children = record_bytes(0, 0, 4026, &utf16("Ada Lovelace"));
        children.extend_from_slice(&record_bytes(
            0,
            1,
            4026,
            &utf16("First\tline\r\nSecond line"),
        ));
        children.extend_from_slice(&record_bytes(0, 2, 4026, &utf16("AL")));
        children.extend_from_slice(&comment_atom(index));
        record_bytes(0x0f, 0, 12000, &children)
    }

    fn author_container(seed: i32) -> Vec<u8> {
        let mut children = record_bytes(0, 0, 4026, &utf16("Ada Lovelace"));
        let mut index = 3i32.to_le_bytes().to_vec();
        index.extend_from_slice(&seed.to_le_bytes());
        children.extend_from_slice(&record_bytes(0, 0, 12005, &index));
        record_bytes(0x0f, 0, 12004, &children)
    }

    #[test]
    fn parses_comments_and_author_metadata() {
        let comment_root = root(vec![prog_tags_record(10, &comment_container(7))]);
        let comments = parse_slide_comments(&comment_root).unwrap();
        assert_eq!(comments.len(), 1);
        let comment = &comments[0];
        assert_eq!(comment.author, "Ada Lovelace");
        assert_eq!(comment.text, "First\tline\r\nSecond line");
        assert_eq!(comment.index, 7);
        assert_eq!(comment.year, 2026);
        assert_eq!(comment.day_of_week, 4);
        assert_eq!(comment.millisecond, 999);
        assert_eq!((comment.x, comment.y), (-12, 34));

        let author_root = root(vec![prog_tags_record(10, &author_container(7))]);
        let authors = PowerPointCommentAuthors::parse(&author_root).unwrap();
        let author = authors.find("Ada Lovelace").unwrap();
        assert_eq!(author.color_index, Some(3));
        assert_eq!(author.comment_index_seed, Some(7));
        authors.validate_comments(&comments).unwrap();
    }

    #[test]
    fn rejects_comment_indices_above_author_seed() {
        let comments =
            parse_slide_comments(&root(vec![prog_tags_record(10, &comment_container(8))])).unwrap();
        let authors = PowerPointCommentAuthors::parse(&root(vec![prog_tags_record(
            10,
            &author_container(7),
        )]))
        .unwrap();

        assert!(authors.validate_comments(&comments).is_err());
    }

    #[test]
    fn ignores_comments_from_other_programmable_tag_versions() {
        let document = root(vec![prog_tags_record(9, &comment_container(7))]);

        assert!(parse_slide_comments(&document).unwrap().is_empty());
        assert!(
            PowerPointCommentAuthors::parse(&document)
                .unwrap()
                .authors
                .is_empty()
        );
    }

    #[test]
    fn rejects_malformed_comment_containers() {
        let mut out_of_order = record_bytes(0, 1, 4026, &utf16("text"));
        out_of_order.extend_from_slice(&record_bytes(0, 0, 4026, &utf16("author")));
        out_of_order.extend_from_slice(&comment_atom(0));
        let mut forbidden_text = record_bytes(0, 1, 4026, &[0x0b, 0]);
        forbidden_text.extend_from_slice(&comment_atom(0));
        let mut duplicate_atom = comment_atom(0);
        duplicate_atom.extend_from_slice(&comment_atom(1));
        let malformed = [
            record_bytes(0x0e, 0, 12000, &comment_atom(0)),
            record_bytes(0x0f, 0, 12000, &[]),
            record_bytes(0x0f, 0, 12000, &comment_atom(-1)),
            record_bytes(0x0f, 0, 12000, &out_of_order),
            record_bytes(0x0f, 0, 12000, &forbidden_text),
            record_bytes(0x0f, 0, 12000, &duplicate_atom),
        ];
        for record in malformed {
            let document = root(vec![prog_tags_record(10, &record)]);
            assert!(parse_slide_comments(&document).is_err());
        }
    }

    #[test]
    fn rejects_malformed_comment_authors() {
        let name = record_bytes(0, 0, 4026, &utf16("Ada Lovelace"));
        let mut negative_color = (-1i32).to_le_bytes().to_vec();
        negative_color.extend_from_slice(&0i32.to_le_bytes());
        let atom = record_bytes(0, 0, 12005, &negative_color);
        let mut out_of_order = atom.clone();
        out_of_order.extend_from_slice(&name);
        let malformed = [
            record_bytes(0x0e, 0, 12004, &name),
            record_bytes(0x0f, 0, 12004, &atom),
            record_bytes(0x0f, 0, 12004, &out_of_order),
            record_bytes(0x0f, 0, 12004, &record_bytes(0, 0, 4026, &[0x01, 0])),
        ];
        for record in malformed {
            let document = root(vec![prog_tags_record(10, &record)]);
            assert!(PowerPointCommentAuthors::parse(&document).is_err());
        }
    }
}
