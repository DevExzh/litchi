//! Lossless discovery of the slide-owned diagram payload boundary.
//!
//! This codec scans record headers and ranges but never re-encodes the
//! enclosing slide. Only the fixed-width `BuildList` bytes are later replaced;
//! all `PPDrawing`, programmable-tag, opaque, and envelope bytes stay in their
//! original positions.

use std::ops::Range;

use crate::consts::RecordType;
use crate::package::Result;

use super::model::SlideLimits;
use super::validation::{corrupted, invalid, validate_limits};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Parts {
    pub(super) build_list: Range<usize>,
    pub(super) drawing: Range<usize>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RecordRange {
    full: Range<usize>,
    data: Range<usize>,
    record_type: u16,
    version: u16,
    instance: u16,
}

/// Locate the one supported diagram owner in a complete `SlideContainer`.
pub(super) fn locate(bytes: &[u8], limits: SlideLimits) -> Result<Parts> {
    validate_limits(limits)?;
    if bytes.len() > limits.max_slide_bytes {
        return Err(invalid(
            "SlideContainer exceeds the configured slide byte limit",
        ));
    }

    let root = read_record(bytes, 0, bytes.len(), "SlideContainer")?;
    if root.full.end != bytes.len() {
        return Err(corrupted(
            "SlideContainer input contains trailing bytes outside its record envelope",
        ));
    }
    if root.record_type != RecordType::Slide.as_u16() || root.version != 0x0F || root.instance != 0
    {
        return Err(invalid(
            "diagram slide publication requires one versioned SlideContainer",
        ));
    }

    let mut drawing = None;
    let mut build_list = None;
    visit_sequence(
        bytes,
        root.data.clone(),
        "SlideContainer children",
        |child| {
            if child.record_type == RecordType::PPDrawing.as_u16() {
                if drawing.is_some() {
                    return Err(invalid(
                        "SlideContainer contains duplicate PPDrawing records",
                    ));
                }
                if child.version != 0x0F || child.instance != 0 {
                    return Err(invalid(
                        "PPDrawing sibling has an unsupported record envelope",
                    ));
                }
                crate::odraw::parse_drawing(&bytes[child.data.clone()])?;
                drawing = Some(child.data);
            } else if child.record_type == RecordType::ProgTags.as_u16() {
                if child.version != 0x0F || child.instance != 0 {
                    return Err(invalid(
                        "SlideProgTagsContainer has an unsupported record envelope",
                    ));
                }
                locate_tags(bytes, child, &mut build_list)?;
            }
            Ok(())
        },
    )?;

    let drawing_range =
        drawing.ok_or_else(|| corrupted("SlideContainer is missing its PPDrawing sibling"))?;
    let build_list_range = build_list
        .ok_or_else(|| invalid("SlideContainer has no supported ___PPT10 BuildList payload"))?;
    Ok(Parts {
        build_list: build_list_range,
        drawing: drawing_range,
    })
}

fn locate_tags(
    bytes: &[u8],
    tags: RecordRange,
    build_list: &mut Option<Range<usize>>,
) -> Result<()> {
    visit_sequence(bytes, tags.data, "ProgTags", |tag| {
        if tag.record_type == RecordType::ProgBinaryTag.as_u16() {
            if tag.version != 0x0F || tag.instance != 0 {
                return Err(invalid(
                    "___PPT10 programmable tag has an unsupported record envelope",
                ));
            }
            locate_binary_tag(bytes, tag, build_list)?;
        }
        Ok(())
    })
}

fn locate_binary_tag(
    bytes: &[u8],
    binary: RecordRange,
    build_list: &mut Option<Range<usize>>,
) -> Result<()> {
    let mut is_ppt10 = false;
    let mut binary_data = None;
    let mut binary_data_count = 0usize;

    visit_sequence(bytes, binary.data, "ProgBinaryTag", |child| {
        if child.record_type == RecordType::CString.as_u16() && is_ppt10_name(bytes, &child) {
            if is_ppt10 {
                return Err(invalid("___PPT10 contains duplicate tag names"));
            }
            is_ppt10 = true;
        } else if child.record_type == RecordType::BinaryTagData.as_u16() {
            if child.version != 0 || child.instance != 0 {
                return Err(invalid(
                    "___PPT10 BinaryTagData has an unsupported record envelope",
                ));
            }
            binary_data_count = binary_data_count
                .checked_add(1)
                .ok_or_else(|| corrupted("___PPT10 BinaryTagData count overflow"))?;
            binary_data = Some(child);
        }
        Ok(())
    })?;

    if !is_ppt10 {
        return Ok(());
    }
    if binary_data_count != 1 {
        return Err(corrupted(
            "___PPT10 must contain exactly one BinaryTagData record",
        ));
    }
    let Some(data) = binary_data else {
        return Err(corrupted(
            "___PPT10 must contain exactly one BinaryTagData record",
        ));
    };
    visit_sequence(bytes, data.data, "___PPT10 BinaryTagData", |child| {
        if child.record_type != RecordType::BuildList.as_u16() {
            return Ok(());
        }
        if child.version != 0x0F || child.instance != 0 {
            return Err(invalid(
                "___PPT10 BuildList has an unsupported record envelope",
            ));
        }
        if build_list.is_some() {
            return Err(invalid(
                "SlideContainer contains duplicate ___PPT10 BuildList records",
            ));
        }
        *build_list = Some(child.full);
        Ok(())
    })
}

fn is_ppt10_name(bytes: &[u8], record: &RecordRange) -> bool {
    if record.version != 0 || record.instance != 0 {
        return false;
    }
    let expected = "___PPT10"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    bytes.get(record.data.clone()) == Some(expected.as_slice())
}

fn visit_sequence(
    bytes: &[u8],
    range: Range<usize>,
    context: &str,
    mut visit: impl FnMut(RecordRange) -> Result<()>,
) -> Result<()> {
    let mut offset = range.start;
    while offset < range.end {
        let record = read_record(bytes, offset, range.end, context)?;
        let end = record.full.end;
        visit(record)?;
        offset = end;
    }
    if offset != range.end {
        return Err(corrupted(format!(
            "{context} records do not cover their payload"
        )));
    }
    Ok(())
}

fn read_record(bytes: &[u8], start: usize, limit: usize, context: &str) -> Result<RecordRange> {
    let header_end = start
        .checked_add(8)
        .ok_or_else(|| corrupted(format!("{context} record header offset overflow")))?;
    if header_end > limit || header_end > bytes.len() {
        return Err(corrupted(format!(
            "{context} has a truncated record header"
        )));
    }
    let packed = u16::from_le_bytes([bytes[start], bytes[start + 1]]);
    let record_type = u16::from_le_bytes([bytes[start + 2], bytes[start + 3]]);
    let data_len = usize::try_from(u32::from_le_bytes([
        bytes[start + 4],
        bytes[start + 5],
        bytes[start + 6],
        bytes[start + 7],
    ]))
    .map_err(|_err| corrupted(format!("{context} record length exceeds this platform")))?;
    let end = header_end
        .checked_add(data_len)
        .ok_or_else(|| corrupted(format!("{context} record length overflows")))?;
    if end > limit || end > bytes.len() {
        return Err(corrupted(format!("{context} record exceeds its envelope")));
    }
    Ok(RecordRange {
        full: start..end,
        data: header_end..end,
        record_type,
        version: packed & 0x000F,
        instance: packed >> 4,
    })
}

/// Replace only the fixed-width `BuildList` range in its owning slide.
pub(super) fn replace_build_list(
    source: &[u8],
    range: Range<usize>,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    if range.start > range.end || range.end > source.len() {
        return Err(corrupted(
            "diagram BuildList publication range is out of bounds",
        ));
    }
    if replacement.len() != range.len() {
        return Err(invalid(
            "diagram publication cannot change BuildList record framing",
        ));
    }
    let mut output = source.to_vec();
    output[range].copy_from_slice(replacement);
    Ok(output)
}
