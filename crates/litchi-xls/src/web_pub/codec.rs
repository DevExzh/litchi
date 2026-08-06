//! BIFF8 `WebPub` payload codec (MS-XLS 2.4.344).

use std::ops::Range;

use crate::{Error, Result};

use super::model::{WebPageType, WebPub, WebPubRange, WebSourceType};
use super::{AUTO_REPUBLISH, FRT_REF, MHTML, WEB_PUB_RECORD_TYPE, invalid};

/// Size in bytes of the fixed record part: `FrtRefHeaderU`, `tws`, `twd`,
/// the flag word, `reserved3`, `unused2`, `nStyleId`, and `cb`.
const FIXED_LEN: usize = 28;
/// Size in bytes of the trailing `unused3` field.
const TRAILING_UNUSED_LEN: usize = 2;
/// Maximum character count of a `WebPubString` (MS-XLS 2.5.278).
const MAX_WEB_PUB_STRING_CHARS: usize = 255;

/// `fHighByte` bit of a BIFF8 string option byte.
const HIGH_BYTE: u8 = 0x01;

/// Wire spans for the fields owned by the semantic `WebPub` model.
///
/// Keeping these spans beside the parsed value lets the edit layer replace a
/// single owned field without normalizing producer-specific string options,
/// reserved bytes, or trailing padding.
#[derive(Debug, Clone)]
pub(super) struct WebPubLayout {
    source_name: Option<Range<usize>>,
    file_destination: Range<usize>,
    div_id: Range<usize>,
    title: Range<usize>,
    chart_shape_id: Option<Range<usize>>,
    reserved: Range<usize>,
    trailing_unused: Range<usize>,
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

impl WebPub {
    /// Parse a `WebPub` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        Self::parse_with_layout(data).map(|(value, _)| value)
    }

    /// Parse a `WebPub` payload and retain the owned/opaque wire spans.
    pub(super) fn parse_with_layout(data: &[u8]) -> Result<(Self, WebPubLayout)> {
        if data.len() < FIXED_LEN + TRAILING_UNUSED_LEN {
            return Err(Error::InvalidLength {
                expected: FIXED_LEN + TRAILING_UNUSED_LEN,
                found: data.len(),
            });
        }
        if read_u16(data, 0) != WEB_PUB_RECORD_TYPE {
            return Err(invalid("WebPub FrtRefHeaderU.rt mismatch"));
        }
        let has_ref = read_u16(data, 2) & FRT_REF != 0;
        let range_ref = (
            read_u16(data, 4),
            read_u16(data, 6),
            read_u16(data, 8),
            read_u16(data, 10),
        );

        let source = WebSourceType::from_code(data[12])?;
        let page_type = WebPageType::from_code(data[13])?;
        let flags = read_u16(data, 14);
        let style_id = read_u32(data, 20);
        let declared_size = read_u32(data, 24);

        // cb counts everything after the fixed part (MS-XLS 2.4.344).
        let tail_size = data.len() - FIXED_LEN;
        if usize::try_from(declared_size) != Ok(tail_size) {
            return Err(invalid("WebPub cb does not match the record size"));
        }
        // Per MS-XLS 2.4.344 the ref8 range applies iff tws is 0x04, and
        // fFrtRef MUST be set exactly in that case.
        if (source == WebSourceType::Range) != has_ref {
            return Err(invalid("WebPub fFrtRef does not match the tws source type"));
        }

        let mut offset = FIXED_LEN;
        let source_name = if source.code() > WebSourceType::Range.code() {
            let start = offset;
            let (name, used) = parse_web_pub_string(&data[offset..])?;
            offset += used;
            Some((name, start..offset))
        } else {
            None
        };
        let start = offset;
        let (file_destination, used) = parse_web_pub_string(&data[offset..])?;
        offset += used;
        let file_destination_span = start..offset;
        let start = offset;
        let (div_id, used) = parse_web_pub_string(&data[offset..])?;
        offset += used;
        let div_id_span = start..offset;
        let start = offset;
        let (title, used) = parse_web_pub_string(&data[offset..])?;
        offset += used;
        let title_span = start..offset;
        let chart_shape_id = if source == WebSourceType::Chart {
            let start = offset;
            let raw = data.get(offset..offset + 4).ok_or(Error::InvalidLength {
                expected: offset + 4,
                found: data.len(),
            })?;
            offset += 4;
            Some((read_u32(raw, 0), start..offset))
        } else {
            None
        };

        // frtRgb fills the record up to the trailing 2-byte unused3 field.
        let reserved_end = data.len() - TRAILING_UNUSED_LEN;
        if offset > reserved_end {
            return Err(invalid("WebPub strings overrun the record"));
        }
        let reserved_span = offset..reserved_end;
        let reserved = data[reserved_span.clone()].to_vec();

        let value = WebPub {
            source,
            page_type,
            range: (source == WebSourceType::Range)
                .then(|| WebPubRange::decode(range_ref.0, range_ref.1, range_ref.2, range_ref.3))
                .transpose()?,
            auto_republish: flags & AUTO_REPUBLISH != 0,
            single_file: flags & MHTML != 0,
            style_id,
            source_name: source_name.as_ref().map(|(name, _)| name.clone()),
            file_destination,
            div_id,
            title,
            chart_shape_id: chart_shape_id.as_ref().map(|(shape_id, _)| *shape_id),
            reserved,
        };
        let layout = WebPubLayout {
            source_name: source_name.map(|(_, span)| span),
            file_destination: file_destination_span,
            div_id: div_id_span,
            title: title_span,
            chart_shape_id: chart_shape_id.map(|(_, span)| span),
            reserved: reserved_span,
            trailing_unused: reserved_end..data.len(),
        };

        Ok((value, layout))
    }

    /// Serialize back to a complete `WebPub` record payload.
    ///
    /// The conditional fields must agree with [`WebPub::source`]:
    /// `range` is required exactly for [`WebSourceType::Range`],
    /// `source_name` is required exactly when the source code is greater
    /// than 4, and `chart_shape_id` exactly for [`WebSourceType::Chart`].
    pub(crate) fn validate_for_write(&self) -> Result<()> {
        let wants_range = self.source == WebSourceType::Range;
        if wants_range != self.range.is_some() {
            return Err(Error::InvalidData(
                "WebPub range must be present iff the source type is Range".to_string(),
            ));
        }
        let wants_name = self.source.code() > WebSourceType::Range.code();
        if wants_name != self.source_name.is_some() {
            return Err(Error::InvalidData(
                "WebPub source_name must be present iff the tws code is greater than 4".to_string(),
            ));
        }
        let wants_shape = self.source == WebSourceType::Chart;
        if wants_shape != self.chart_shape_id.is_some() {
            return Err(Error::InvalidData(
                "WebPub chart_shape_id must be present iff the source type is Chart".to_string(),
            ));
        }
        for text in [
            self.source_name.as_deref(),
            Some(self.file_destination.as_str()),
            Some(self.div_id.as_str()),
            Some(self.title.as_str()),
        ]
        .into_iter()
        .flatten()
        {
            web_pub_string_layout(text)?;
        }
        Ok(())
    }

    pub(crate) fn to_payload(&self) -> Result<Vec<u8>> {
        self.validate_for_write()?;
        let wants_range = self.source == WebSourceType::Range;

        let mut tail = Vec::new();
        if let Some(name) = &self.source_name {
            write_web_pub_string(&mut tail, name)?;
        }
        write_web_pub_string(&mut tail, &self.file_destination)?;
        write_web_pub_string(&mut tail, &self.div_id)?;
        write_web_pub_string(&mut tail, &self.title)?;
        if let Some(shape_id) = self.chart_shape_id {
            tail.extend_from_slice(&shape_id.to_le_bytes());
        }
        tail.extend_from_slice(&self.reserved);
        tail.extend_from_slice(&[0u8; TRAILING_UNUSED_LEN]); // unused3

        let mut payload = Vec::with_capacity(FIXED_LEN + tail.len());
        payload.extend_from_slice(&WEB_PUB_RECORD_TYPE.to_le_bytes());
        let mut grbit_frt = 0u16;
        if wants_range {
            grbit_frt |= FRT_REF;
        }
        payload.extend_from_slice(&grbit_frt.to_le_bytes());
        let (first_row, last_row, first_column, last_column) =
            self.range.map_or((0, 0, 0, 0), WebPubRange::fields);
        payload.extend_from_slice(&first_row.to_le_bytes());
        payload.extend_from_slice(&last_row.to_le_bytes());
        payload.extend_from_slice(&u16::from(first_column).to_le_bytes());
        payload.extend_from_slice(&u16::from(last_column).to_le_bytes());
        payload.push(self.source.code());
        payload.push(self.page_type.code());
        let mut flags = 0u16;
        if self.auto_republish {
            flags |= AUTO_REPUBLISH;
        }
        if self.single_file {
            flags |= MHTML;
        }
        payload.extend_from_slice(&flags.to_le_bytes());
        payload.extend_from_slice(&[0u8; 4]); // reserved3 + unused2
        payload.extend_from_slice(&self.style_id.to_le_bytes());
        payload.extend_from_slice(&(tail.len() as u32).to_le_bytes());
        payload.extend_from_slice(&tail);
        Ok(payload)
    }
}

/// Rewrite only semantic `WebPub` fields while retaining all source-owned
/// bytes outside the requested fields.
pub(super) fn rewrite_preserving_source(
    source: &[u8],
    current: &WebPub,
    replacement: &WebPub,
) -> Result<Vec<u8>> {
    let (parsed, layout) = WebPub::parse_with_layout(source)?;
    if &parsed != current {
        return Err(Error::UnsafeEdit(
            "WebPub transaction candidate is out of sync with its source bytes".into(),
        ));
    }
    replacement.validate_for_write()?;
    if replacement.source != current.source
        || replacement.range.is_some() != current.range.is_some()
        || replacement.source_name.is_some() != current.source_name.is_some()
        || replacement.chart_shape_id.is_some() != current.chart_shape_id.is_some()
    {
        return Err(Error::UnsafeEdit(
            "WebPub edits cannot change conditional record topology".into(),
        ));
    }

    let mut tail = Vec::new();
    if let Some(span) = &layout.source_name {
        let before = current.source_name.as_deref().ok_or_else(|| {
            Error::UnsafeEdit("WebPub source-name span has no semantic owner".into())
        })?;
        let after = replacement
            .source_name
            .as_deref()
            .ok_or_else(|| Error::UnsafeEdit("WebPub source-name replacement is missing".into()))?;
        append_string(&mut tail, source, span, before, after)?;
    }
    append_string(
        &mut tail,
        source,
        &layout.file_destination,
        &current.file_destination,
        &replacement.file_destination,
    )?;
    append_string(
        &mut tail,
        source,
        &layout.div_id,
        &current.div_id,
        &replacement.div_id,
    )?;
    append_string(
        &mut tail,
        source,
        &layout.title,
        &current.title,
        &replacement.title,
    )?;

    if let Some(span) = &layout.chart_shape_id {
        let before = current.chart_shape_id.ok_or_else(|| {
            Error::UnsafeEdit("WebPub chart-shape span has no semantic owner".into())
        })?;
        let after = replacement
            .chart_shape_id
            .ok_or_else(|| Error::UnsafeEdit("WebPub chart-shape replacement is missing".into()))?;
        if before == after {
            tail.extend_from_slice(&source[span.clone()]);
        } else {
            tail.extend_from_slice(&after.to_le_bytes());
        }
    }
    tail.extend_from_slice(&source[layout.reserved.clone()]);
    tail.extend_from_slice(&source[layout.trailing_unused.clone()]);

    let tail_len = u32::try_from(tail.len())
        .map_err(|_| Error::InvalidData("WebPub payload tail exceeds u32".into()))?;
    let mut payload = source[..FIXED_LEN].to_vec();
    if replacement.range != current.range {
        let (first_row, last_row, first_column, last_column) =
            replacement.range.map_or((0, 0, 0, 0), WebPubRange::fields);
        payload[4..6].copy_from_slice(&first_row.to_le_bytes());
        payload[6..8].copy_from_slice(&last_row.to_le_bytes());
        payload[8..10].copy_from_slice(&u16::from(first_column).to_le_bytes());
        payload[10..12].copy_from_slice(&u16::from(last_column).to_le_bytes());
    }
    if replacement.page_type != current.page_type {
        payload[13] = replacement.page_type.code();
    }
    if replacement.auto_republish != current.auto_republish
        || replacement.single_file != current.single_file
    {
        let mut flags = read_u16(source, 14) & !(AUTO_REPUBLISH | MHTML);
        if replacement.auto_republish {
            flags |= AUTO_REPUBLISH;
        }
        if replacement.single_file {
            flags |= MHTML;
        }
        payload[14..16].copy_from_slice(&flags.to_le_bytes());
    }
    if replacement.style_id != current.style_id {
        payload[20..24].copy_from_slice(&replacement.style_id.to_le_bytes());
    }
    payload[24..28].copy_from_slice(&tail_len.to_le_bytes());
    payload.extend_from_slice(&tail);
    Ok(payload)
}

fn append_string(
    out: &mut Vec<u8>,
    source: &[u8],
    span: &Range<usize>,
    before: &str,
    after: &str,
) -> Result<()> {
    if before == after {
        out.extend_from_slice(&source[span.clone()]);
    } else {
        write_web_pub_string(out, after)?;
    }
    Ok(())
}

/// Serialize a `WebPubString` (MS-XLS 2.5.278), compressed when every
/// character is in U+0000..=U+00FF and wide otherwise.
fn write_web_pub_string(out: &mut Vec<u8>, text: &str) -> Result<()> {
    let (compressible, char_count) = web_pub_string_layout(text)?;
    out.extend_from_slice(&(char_count as u16).to_le_bytes());
    if compressible {
        out.push(0u8); // fHighByte = 0
        out.extend(text.chars().map(|ch| ch as u8));
    } else {
        out.push(HIGH_BYTE);
        for unit in text.encode_utf16() {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    }
    Ok(())
}

fn web_pub_string_layout(text: &str) -> Result<(bool, usize)> {
    let compressible = text.chars().all(|ch| u32::from(ch) <= 0xFF);
    let char_count = if compressible {
        text.chars().count()
    } else {
        text.encode_utf16().count()
    };
    if char_count > MAX_WEB_PUB_STRING_CHARS {
        return Err(Error::InvalidData(
            "WebPubString exceeds 255 characters".to_string(),
        ));
    }
    Ok((compressible, char_count))
}

/// Parse a `WebPubString` (MS-XLS 2.5.278): a 2-byte character count
/// followed by an `XLUnicodeStringNoCch`. Returns the string and the number
/// of bytes consumed.
fn parse_web_pub_string(data: &[u8]) -> Result<(String, usize)> {
    if data.len() < 3 {
        return Err(Error::InvalidLength {
            expected: 3,
            found: data.len(),
        });
    }
    let char_count = usize::from(read_u16(data, 0));
    if char_count > MAX_WEB_PUB_STRING_CHARS {
        return Err(invalid("WebPubString exceeds 255 characters"));
    }
    let wide = data[2] & HIGH_BYTE != 0;
    let byte_len = if wide { char_count * 2 } else { char_count };
    let bytes = data.get(3..3 + byte_len).ok_or(Error::InvalidLength {
        expected: 3 + byte_len,
        found: data.len(),
    })?;
    let text = if wide {
        let units: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        String::from_utf16(&units).map_err(|_| invalid("WebPubString is not valid UTF-16LE"))?
    } else {
        bytes.iter().map(|&byte| char::from(byte)).collect()
    };
    Ok((text, 3 + byte_len))
}
