use super::{DateTime, WriteError};

/// Bytes summed per `u8` accumulator by [`utf16_units`]. Each byte adds at most
/// two code units, so a block adds at most 254 and the sum cannot wrap.
const UNIT_COUNT_BLOCK: usize = 127;

/// Returns how many UTF-16 code units `text` encodes to, without decoding it.
///
/// ASCII text is one unit per byte. Otherwise every scalar value contributes
/// one unit for its leading byte (each byte that is not a `0b10xx_xxxx`
/// continuation byte) and a second unit when that byte starts a four-byte
/// sequence (`0xF0..=0xF4`): exactly the supplementary-plane scalar values,
/// which UTF-16 encodes as a surrogate pair. A `&str` is valid UTF-8, so the
/// result always equals `text.encode_utf16().count()`.
pub(in crate::writer::core) fn utf16_units(text: &str) -> usize {
    if text.is_ascii() {
        return text.len();
    }
    text.as_bytes()
        .chunks(UNIT_COUNT_BLOCK)
        .map(|block| {
            usize::from(block.iter().fold(0u8, |units, &byte| {
                units.wrapping_add(u8::from(byte & 0xC0 != 0x80) + u8::from(byte >= 0xF0))
            }))
        })
        .sum()
}

/// Returns the UTF-16 code-unit length of `text` as a checked MS-DOC CP count.
pub(in crate::writer::core) fn utf16_code_unit_len(text: &str) -> Result<u32, WriteError> {
    let length = u32::try_from(utf16_units(text))
        .map_err(|_| WriteError::InvalidData("DOC text exceeds the 32-bit CP range".to_string()))?;
    if length >= 0x7FFF_FFFF {
        return Err(WriteError::InvalidData(
            "DOC text exceeds the MS-DOC CP limit".to_string(),
        ));
    }
    Ok(length)
}

/// The `WordDocument` stream while the stories append their text: the zeroed
/// FIB placeholder and its padding, then the text.
///
/// The stories see only the text. [`Self::text_len`] and [`Self::text`]
/// exclude the placeholder, so every FC a story computes as
/// `text_fc_start + text_len()` and every text-relative offset is what it was
/// when the text had a buffer of its own; the finished stream simply keeps the
/// text where it was written instead of copying it behind the placeholder.
pub(in crate::writer::core) struct TextStream {
    stream: Vec<u8>,
    text_start: usize,
}

impl TextStream {
    /// Starts a stream with a `text_start`-byte zeroed prefix and room for at
    /// least `capacity` bytes in all.
    pub(in crate::writer::core) fn new(text_start: usize, capacity: usize) -> Self {
        let mut stream = Vec::with_capacity(capacity.max(text_start));
        stream.resize(text_start, 0);
        Self { stream, text_start }
    }

    /// Bytes of text appended so far, excluding the prefix.
    pub(in crate::writer::core) fn text_len(&self) -> usize {
        self.stream.len() - self.text_start
    }

    /// The text appended so far, excluding the prefix.
    pub(in crate::writer::core) fn text(&self) -> &[u8] {
        &self.stream[self.text_start..]
    }

    /// Appends raw text-stream bytes, such as a paragraph mark.
    pub(in crate::writer::core) fn extend_from_slice(&mut self, bytes: &[u8]) {
        self.stream.extend_from_slice(bytes);
    }

    /// Appends all of `text` as UTF-16LE.
    ///
    /// `units` is the length [`utf16_code_unit_len`] returned for `text`; it
    /// only sizes the single reservation, so the stream grows at most once by
    /// the encoded size. ASCII text is widened byte by byte without decoding;
    /// other text is decoded and encoded in one pass.
    pub(in crate::writer::core) fn push_utf16le(&mut self, text: &str, units: u32) {
        // A checked length is below 2^31, so the encoded byte count fits `usize`.
        self.stream.reserve(units as usize * 2);
        if text.is_ascii() {
            self.stream.extend(text.bytes().flat_map(|byte| [byte, 0]));
        } else {
            for unit in text.encode_utf16() {
                self.stream.extend_from_slice(&unit.to_le_bytes());
            }
        }
    }

    /// Checks the length of `text` exactly as [`utf16_code_unit_len`] does,
    /// then appends it as UTF-16LE and returns that length.
    ///
    /// Nothing is appended when the length is refused.
    pub(in crate::writer::core) fn append_utf16le(
        &mut self,
        text: &str,
    ) -> Result<u32, WriteError> {
        let units = utf16_code_unit_len(text)?;
        self.push_utf16le(text, units);
        Ok(units)
    }

    /// The finished stream: the prefix followed by the text.
    pub(in crate::writer::core) fn into_stream(self) -> Vec<u8> {
        self.stream
    }
}

/// Returns whether `text` contains a field begin, separator, or end character
/// (U+0013, U+0014, U+0015).
///
/// These are ASCII control characters, and a UTF-8 byte below 0x80 is always
/// the complete encoding of an ASCII character, so the byte scan is exact.
/// Each block is tested without an early exit so the scan vectorizes.
pub(in crate::writer::core) fn contains_field_character(text: &str) -> bool {
    text.as_bytes().chunks(256).any(|block| {
        block
            .iter()
            .fold(false, |found, &byte| found | (byte.wrapping_sub(0x13) < 3))
    })
}

pub(crate) fn pack_dttm(value: Option<DateTime>) -> Result<u32, WriteError> {
    let Some(value) = value else {
        return Ok(0);
    };
    if !(1900..=2411).contains(&value.year)
        || !(1..=12).contains(&value.month)
        || !(1..=31).contains(&value.day)
        || value.hour > 23
        || value.minute > 59
        || value.weekday > 6
    {
        return Err(WriteError::InvalidData(
            "DOC timestamp is outside the DTTM field ranges".to_string(),
        ));
    }
    Ok(u32::from(value.minute)
        | (u32::from(value.hour) << 6)
        | (u32::from(value.day) << 11)
        | (u32::from(value.month) << 16)
        | (u32::from(value.year - 1900) << 20)
        | (u32::from(value.weekday) << 29))
}
