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

/// The largest up-front reservation of the `WordDocument` stream: the 4 GiB its
/// 32-bit FCs address. A larger estimate only belongs to a document whose
/// stream grows past the reservation, as streams always could, or that its FC
/// checks refuse.
const MAX_WORD_DOCUMENT_RESERVATION: usize = u32::MAX as usize;

impl TextStream {
    /// Starts a stream with a `text_start`-byte zeroed prefix and room for at
    /// least `capacity` bytes in all, up to [`MAX_WORD_DOCUMENT_RESERVATION`].
    ///
    /// # Errors
    ///
    /// The reservation is fallible: one that cannot be made is refused with a
    /// typed error rather than aborting the process.
    pub(in crate::writer::core) fn try_new(
        text_start: usize,
        capacity: usize,
    ) -> Result<Self, WriteError> {
        let mut stream =
            reserved_bytes(capacity.max(text_start).min(MAX_WORD_DOCUMENT_RESERVATION))?;
        stream.resize(text_start, 0);
        Ok(Self { stream, text_start })
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

/// A new buffer with room for exactly `bytes`, or a typed refusal when that
/// much cannot be reserved.
fn reserved_bytes(bytes: usize) -> Result<Vec<u8>, WriteError> {
    let mut buffer = Vec::new();
    buffer.try_reserve_exact(bytes).map_err(|_| {
        WriteError::InvalidData("DOC WordDocument stream allocation is too large".to_string())
    })?;
    Ok(buffer)
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

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    const FIELD_CHARACTERS: std::ops::RangeInclusive<char> = '\u{13}'..='\u{15}';

    /// Every Unicode scalar value in order: 1,112,064 of them in 4,382,592
    /// UTF-8 bytes, with every UTF-8 lead and continuation byte value.
    fn every_scalar() -> String {
        (0..=0x10_FFFF_u32).filter_map(char::from_u32).collect()
    }

    fn padded(padding: usize, text: &str) -> String {
        let mut padded = "x".repeat(padding);
        padded.push_str(text);
        padded
    }

    /// Shifting the text by 0 to 63 bytes puts every scalar at every offset
    /// from a 64-byte boundary, straddling one wherever its bytes allow.
    #[test]
    fn utf16_units_counts_every_scalar_at_every_offset_from_a_64_byte_boundary() {
        let every = every_scalar();
        assert_eq!(every.len(), 4_382_592);
        let expected = every.encode_utf16().count();
        assert_eq!(expected, 2_160_640);
        for shift in 0..64 {
            assert_eq!(
                utf16_units(&padded(shift, &every)),
                expected + shift,
                "shift {shift}"
            );
        }
    }

    /// The count sums `UNIT_COUNT_BLOCK`-byte blocks: each scalar length, split
    /// every possible way across the first three block boundaries.
    #[test]
    fn utf16_units_counts_scalars_split_across_its_blocks() {
        for scalar in [
            '\u{80}',
            '\u{7FF}',
            '\u{800}',
            '\u{D7FF}',
            '\u{E000}',
            '\u{FFFF}',
            '\u{10000}',
            '\u{10FFFF}',
        ] {
            for boundary in [UNIT_COUNT_BLOCK, 2 * UNIT_COUNT_BLOCK, 3 * UNIT_COUNT_BLOCK] {
                for before in 0..=scalar.len_utf8() {
                    let mut text = padded(boundary - before, &scalar.to_string());
                    text.push_str("tail é 😀");
                    assert_eq!(
                        utf16_units(&text),
                        text.encode_utf16().count(),
                        "{scalar:?} with {before} bytes before byte {boundary}"
                    );
                }
            }
        }
        // Dense four-byte scalars: the most code units per byte a block adds
        // outside the ASCII fast path.
        assert_eq!(utf16_units(&"😀".repeat(1_000)), 2_000);
        for text in ["", "ascii only", "é", &"é".repeat(1_000)] {
            assert_eq!(utf16_units(text), text.encode_utf16().count());
        }
    }

    /// No scalar other than U+0013 through U+0015 is a field character, at
    /// any offset from a 64-byte boundary.
    #[test]
    fn contains_field_character_ignores_every_other_scalar_at_every_offset() {
        let others: String = every_scalar()
            .chars()
            .filter(|character| !FIELD_CHARACTERS.contains(character))
            .collect();
        for shift in 0..64 {
            assert!(
                !contains_field_character(&padded(shift, &others)),
                "shift {shift}"
            );
        }
    }

    /// Each field character is found at every byte offset through the second
    /// 256-byte block boundary, and after every scalar of a mixed-width text.
    #[test]
    fn contains_field_character_finds_each_field_character_at_every_offset() {
        let neighbours = "aé漢😀".repeat(100);
        assert!(!contains_field_character(&neighbours));
        for field in FIELD_CHARACTERS {
            for offset in 0..=600 {
                let text = padded(offset, &format!("{field}{neighbours}"));
                assert!(contains_field_character(&text), "{field:?} at {offset}");
            }
            for (index, _) in neighbours.char_indices() {
                let text = format!("{}{field}{}", &neighbours[..index], &neighbours[index..]);
                assert!(contains_field_character(&text), "{field:?} at byte {index}");
            }
        }
    }

    #[test]
    fn a_text_stream_reserves_its_estimate_behind_the_zeroed_prefix() {
        let stream = TextStream::try_new(1_536, 10_000).unwrap();
        assert_eq!(stream.text_len(), 0);
        assert!(stream.stream.capacity() >= 10_000);
        assert_eq!(stream.into_stream(), vec![0; 1_536]);
        // An estimate below the prefix still holds the prefix.
        let small = TextStream::try_new(1_536, 0).unwrap();
        assert_eq!(small.into_stream().len(), 1_536);
    }

    #[test]
    fn a_reservation_that_cannot_be_made_is_refused_not_aborted() {
        assert!(matches!(
            reserved_bytes(usize::MAX),
            Err(WriteError::InvalidData(message)) if message.contains("allocation is too large")
        ));
        assert!(reserved_bytes(0).unwrap().is_empty());
    }
}
