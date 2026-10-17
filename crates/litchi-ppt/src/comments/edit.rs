//! Record-level comment edits of one `SlideContainer` (record 0323).
//!
//! `PowerPoint` 2002 stores a slide's comments as `Comment10Container`
//! records in the `BinaryTagDataBlob` of the slide's `___PPT10` programmable
//! tag ([MS-PPT] 2.3.? `PP10SlideBinaryTagExtension`). An edit rewrites only
//! that path: the selected comment, the blob, its tag and the slide's
//! `SlideProgTagsContainer`. A slide without the tag gains one, placed after
//! the slide records [MS-PPT] 2.5.1 orders before `slideProgTagsContainer`.

use crate::binary_tag_edit::{self, Blob};
use crate::consts::RecordType;
use crate::package::Error;
use crate::record_wire::{self as wire, HEADER_LEN};
use crate::slide_order::{Error as EditError, Refusal, Result};

const ELEMENT: &str = "comment";

fn unsupported(reason: &'static str) -> EditError {
    EditError::Refused(Refusal::UnsupportedElement {
        element: ELEMENT,
        reason,
    })
}

fn not_found() -> EditError {
    EditError::Refused(Refusal::ElementNotFound { element: ELEMENT })
}

const CSTRING: u16 = RecordType::CString as u16;
const COMMENT: u16 = RecordType::Comment2000 as u16;
const COMMENT_ATOM: u16 = RecordType::Comment2000Atom as u16;
const TEXT_MASTER_STYLE10: u16 = RecordType::TextMasterStyle10Atom as u16;
const AUTHOR_MAX_UNITS: usize = 52;
const AUTHOR_TOO_LONG: &str = "the author exceeds 52 UTF-16 units";
const INITIALS_TOO_LONG: &str = "the initials exceed 52 UTF-16 units";
const TEXT_TOO_LONG: &str = "the text exceeds 32,000 UTF-16 units";
const TEXT_MAX_UNITS: usize = 32_000;

/// The slide records [MS-PPT] 2.5.1 places before `slideProgTagsContainer`.
const BEFORE_PROG_TAGS: [u16; 7] = [
    RecordType::SlideAtom as u16,
    RecordType::SSSlideInfoAtom as u16,
    RecordType::HeadersFooters as u16,
    RecordType::RoundTripSlideSyncInfo12 as u16,
    RecordType::PPDrawing as u16,
    0x07F0,  // SlideSchemeColorSchemeAtom
    CSTRING, // SlideNameAtom
];

/// A malformed comment date.
struct Invalid;

/// A comment's creation time, as the `SYSTEMTIME` of a `Comment10Atom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DateTime {
    pub year: u16,
    pub month: u16,
    pub day: u16,
    pub hour: u16,
    pub minute: u16,
    pub second: u16,
    pub millisecond: u16,
}

impl DateTime {
    /// Parse `YYYY-MM-DD`, optionally followed by `THH:MM`, `:SS`, `.fff`
    /// and a trailing `Z`.
    ///
    /// Returns `None` for any other form or an out-of-range field.
    #[must_use]
    pub fn parse(value: &str) -> Option<Self> {
        Self::parse_checked(value).ok()
    }

    fn parse_checked(value: &str) -> std::result::Result<Self, Invalid> {
        let invalid = || Invalid;
        let value = value.strip_suffix('Z').unwrap_or(value);
        let (date, time) = value.split_once('T').unwrap_or((value, ""));
        let number = |text: &str, digits: usize| -> std::result::Result<u16, Invalid> {
            if text.len() != digits || !text.bytes().all(|byte| byte.is_ascii_digit()) {
                return Err(invalid());
            }
            text.parse().map_err(|_error| invalid())
        };
        let mut parts = date.split('-');
        let (Some(year), Some(month), Some(day), None) =
            (parts.next(), parts.next(), parts.next(), parts.next())
        else {
            return Err(invalid());
        };
        let mut result = Self {
            year: number(year, 4)?,
            month: number(month, 2)?,
            day: number(day, 2)?,
            ..Self::default()
        };
        if !time.is_empty() {
            let (clock, fraction) = time.split_once('.').unwrap_or((time, ""));
            let mut fields = clock.split(':');
            result.hour = number(fields.next().ok_or_else(invalid)?, 2)?;
            result.minute = number(fields.next().ok_or_else(invalid)?, 2)?;
            if let Some(second) = fields.next() {
                result.second = number(second, 2)?;
            }
            if fields.next().is_some() {
                return Err(invalid());
            }
            if !fraction.is_empty() {
                if fraction.len() > 3 || !fraction.bytes().all(|byte| byte.is_ascii_digit()) {
                    return Err(invalid());
                }
                let scaled = format!("{fraction:0<3}");
                result.millisecond = scaled.parse().map_err(|_error| invalid())?;
            }
        }
        if !(1601..=30827).contains(&result.year)
            || !(1..=12).contains(&result.month)
            || result.day == 0
            || result.day > days_in_month(result.year, result.month)
            || result.hour > 23
            || result.minute > 59
            || result.second > 59
        {
            return Err(invalid());
        }
        Ok(result)
    }

    /// The `SYSTEMTIME` day of the week (0 = Sunday).
    #[must_use]
    pub fn day_of_week(self) -> u16 {
        // Sakamoto's method over the proleptic Gregorian calendar.
        const OFFSETS: [u16; 12] = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4];
        let year = if self.month < 3 {
            self.year - 1
        } else {
            self.year
        };
        let month = usize::from(self.month.clamp(1, 12)) - 1;
        (year + year / 4 - year / 100 + year / 400 + OFFSETS[month] + self.day) % 7
    }
}

impl std::fmt::Display for DateTime {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.{:03}",
            self.year, self.month, self.day, self.hour, self.minute, self.second, self.millisecond
        )
    }
}

const fn days_in_month(year: u16, month: u16) -> u16 {
    match month {
        4 | 6 | 9 | 11 => 30,
        2 if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 => 29,
        2 => 28,
        _ => 31,
    }
}

/// A comment to append to a slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewComment {
    /// The author's display name (at most 52 UTF-16 units).
    pub author: String,
    /// The author's initials (at most 52 UTF-16 units).
    pub initials: String,
    /// The comment text (at most 32,000 UTF-16 units).
    pub text: String,
    /// The creation time; `None` writes an all-zero `SYSTEMTIME`, since a
    /// writer never consults the ambient clock (ADR 0034 §9).
    pub date: Option<DateTime>,
    /// The anchor, in master units (1/576 inch).
    pub x: i32,
    /// The anchor, in master units.
    pub y: i32,
}

/// A change to an existing comment; `None` keeps a field.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CommentChange {
    pub text: Option<String>,
    pub x: Option<i32>,
    pub y: Option<i32>,
}

fn check_units(text: &str, maximum: usize, reason: &'static str) -> Result<()> {
    if text.encode_utf16().count() > maximum {
        return Err(EditError::Refused(Refusal::InvalidElementValue {
            element: ELEMENT,
            reason,
        }));
    }
    Ok(())
}

fn atom_payload(index: i32, date: Option<DateTime>, x: i32, y: i32) -> [u8; 28] {
    let mut payload = [0_u8; 28];
    payload[0..4].copy_from_slice(&index.to_le_bytes());
    if let Some(date) = date {
        let fields = [
            date.year,
            date.month,
            date.day_of_week(),
            date.day,
            date.hour,
            date.minute,
            date.second,
            date.millisecond,
        ];
        for (slot, value) in fields.into_iter().enumerate() {
            payload[4 + slot * 2..6 + slot * 2].copy_from_slice(&value.to_le_bytes());
        }
    }
    payload[20..24].copy_from_slice(&x.to_le_bytes());
    payload[24..28].copy_from_slice(&y.to_le_bytes());
    payload
}

/// Append `comment` with the per-author `index`; returns the new slide
/// record and the comment's zero-based position among the slide's comments.
pub(crate) fn add(slide: &[u8], comment: &NewComment, index: i32) -> Result<(Vec<u8>, usize)> {
    check_units(&comment.author, AUTHOR_MAX_UNITS, AUTHOR_TOO_LONG)?;
    check_units(&comment.initials, AUTHOR_MAX_UNITS, INITIALS_TOO_LONG)?;
    check_units(&comment.text, TEXT_MAX_UNITS, TEXT_TOO_LONG)?;
    let located = locate(slide)?;
    let comments = comment_indexes(&located)?;
    let record = wire::container(
        COMMENT,
        0,
        &[
            wire::encode(CSTRING, 0, 0, &wire::utf16(&comment.author))?,
            wire::encode(CSTRING, 0, 1, &wire::utf16(&comment.text))?,
            wire::encode(CSTRING, 0, 2, &wire::utf16(&comment.initials))?,
            wire::encode(
                COMMENT_ATOM,
                0,
                0,
                &atom_payload(index, comment.date, comment.x, comment.y),
            )?,
        ],
    )?;
    let insert_at = match comments.last() {
        Some(last) => last + 1,
        None => {
            let mut position = 0;
            for (offset, record) in located.records().iter().enumerate() {
                if wire::record_type(record)? == TEXT_MASTER_STYLE10 {
                    position = offset + 1;
                }
            }
            position
        },
    };
    let mut records = located
        .records()
        .iter()
        .map(|record| record.to_vec())
        .collect::<Vec<_>>();
    records.insert(insert_at, record);
    Ok((
        located.rebuild(slide, &records, &BEFORE_PROG_TAGS)?,
        comments.len(),
    ))
}

/// Apply `change` to the comment at `position`; `None` when nothing changes.
pub(crate) fn set(
    slide: &[u8],
    position: usize,
    change: &CommentChange,
) -> Result<Option<Vec<u8>>> {
    if let Some(text) = &change.text {
        check_units(text, TEXT_MAX_UNITS, TEXT_TOO_LONG)?;
    }
    let located = locate(slide)?;
    let comments = comment_indexes(&located)?;
    let index = *comments.get(position).ok_or_else(not_found)?;
    let comment = located.records()[index];
    let mut changed = false;
    let mut children = Vec::new();
    let mut has_text = false;
    for child in wire::children(comment)? {
        let child = child?;
        let kind = wire::record_type(child)?;
        let instance = wire::record_instance(child)?;
        if kind == CSTRING && instance == 1 {
            has_text = true;
            if let Some(text) = &change.text {
                let units = wire::utf16(text);
                if wire::payload(child)? != units.as_slice() {
                    changed = true;
                    children.push(wire::rebuild(child, &units)?);
                    continue;
                }
            }
        } else if kind == COMMENT_ATOM {
            if let Some(text) = change.text.as_ref().filter(|_| !has_text) {
                // A comment without a text atom gains one before its atom.
                changed = true;
                children.push(wire::encode(CSTRING, 0, 1, &wire::utf16(text))?);
            }
            let mut payload = wire::payload(child)?.to_vec();
            if payload.len() != 28 {
                return Err(Error::Corrupted("Comment10Atom has an invalid size".into()).into());
            }
            if let Some(x) = change.x {
                payload[20..24].copy_from_slice(&x.to_le_bytes());
            }
            if let Some(y) = change.y {
                payload[24..28].copy_from_slice(&y.to_le_bytes());
            }
            if payload.as_slice() != wire::payload(child)? {
                changed = true;
                children.push(wire::rebuild(child, &payload)?);
                continue;
            }
        }
        children.push(child.to_vec());
    }
    if !changed {
        return Ok(None);
    }
    let mut records = located
        .records()
        .iter()
        .map(|record| record.to_vec())
        .collect::<Vec<_>>();
    records[index] = wire::rebuild(comment, &children.concat())?;
    located
        .rebuild(slide, &records, &BEFORE_PROG_TAGS)
        .map(Some)
}

/// Remove the comment at `position`.
pub(crate) fn remove(slide: &[u8], position: usize) -> Result<Vec<u8>> {
    let located = locate(slide)?;
    let comments = comment_indexes(&located)?;
    let index = *comments.get(position).ok_or_else(not_found)?;
    let records = located
        .records()
        .iter()
        .enumerate()
        .filter(|(offset, _)| *offset != index)
        .map(|(_, record)| record.to_vec())
        .collect::<Vec<_>>();
    located.rebuild(slide, &records, &BEFORE_PROG_TAGS)
}

/// Raise the `commentIndexSeed` of the document-level author named
/// `author` to `index`; `None` when the document declares no such author or
/// its seed is already at least `index`.
pub(crate) fn raise_author_seed(
    document: &[u8],
    author: &str,
    index: i32,
) -> Result<Option<Vec<u8>>> {
    if wire::record_type(document)? != RecordType::Document as u16 {
        return Err(
            Error::Corrupted("comment authors owner is not a DocumentContainer".into()).into(),
        );
    }
    // [MS-PPT] 2.4.1: the document's programmable tags live in its
    // `DocInfoListContainer`.
    let children = wire::child_list(document)?;
    let mut info = None;
    for (position, child) in children.iter().enumerate() {
        if wire::record_type(child)? == RecordType::DocInfoList as u16
            && info.replace(position).is_some()
        {
            return Err(unsupported("the document has several information lists"));
        }
    }
    let Some(info) = info else {
        return Ok(None);
    };
    let Some(rewritten) = raise_seed_in(children[info], author, index)? else {
        return Ok(None);
    };
    let mut data = Vec::with_capacity(document.len().saturating_add(8));
    for (position, child) in children.iter().enumerate() {
        if position == info {
            data.extend_from_slice(&rewritten);
        } else {
            data.extend_from_slice(child);
        }
    }
    Ok(Some(wire::rebuild(document, &data)?))
}

fn raise_seed_in(list: &[u8], author: &str, index: i32) -> Result<Option<Vec<u8>>> {
    let located =
        binary_tag_edit::locate(list, RecordType::DocInfoList as u16, "___PPT10", ELEMENT)?;
    let document = list;
    let name = wire::utf16(author);
    let mut records = Vec::with_capacity(located.records().len());
    let mut changed = false;
    for record in located.records() {
        if wire::record_type(record)? != RecordType::CommentIndex10 as u16 || changed {
            records.push(record.to_vec());
            continue;
        }
        let children = wire::child_list(record)?;
        let named = children.first().is_some_and(|child| {
            wire::record_type(child).ok() == Some(CSTRING)
                && wire::payload(child).ok() == Some(name.as_slice())
        });
        let mut rewritten = Vec::with_capacity(children.len());
        for child in &children {
            if named && wire::record_type(child)? == RecordType::CommentIndex10Atom as u16 {
                let mut payload = wire::payload(child)?.to_vec();
                if payload.len() != 8 {
                    return Err(
                        Error::Corrupted("CommentIndex10Atom has an invalid size".into()).into(),
                    );
                }
                let seed = i32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]]);
                if seed < index {
                    payload[4..8].copy_from_slice(&index.to_le_bytes());
                    changed = true;
                    rewritten.push(wire::rebuild(child, &payload)?);
                    continue;
                }
            }
            rewritten.push(child.to_vec());
        }
        records.push(if changed {
            wire::rebuild(record, &rewritten.concat())?
        } else {
            record.to_vec()
        });
    }
    if !changed {
        return Ok(None);
    }
    located.rebuild(document, &records, &[]).map(Some)
}

fn comment_indexes(blob: &Blob<'_>) -> Result<Vec<usize>> {
    let mut indexes = Vec::new();
    for (index, record) in blob.records().iter().enumerate() {
        if wire::record_type(record)? == COMMENT {
            indexes.push(index);
        }
    }
    Ok(indexes)
}

fn locate(slide: &[u8]) -> Result<Blob<'_>> {
    binary_tag_edit::locate(slide, RecordType::Slide as u16, "___PPT10", ELEMENT)
}

const _: () = assert!(HEADER_LEN == 8);
