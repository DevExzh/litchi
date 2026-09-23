//! Shared String Table (SST) BIFF8 writer.

use crate::Result;
use crate::writer::string_limits::{SHARED_STRING_UNITS, ensure_utf16_len_within, u16_len};
use std::io::Write;

use super::write_record_header;

/// Write CONTINUE record
///
/// Record type: 0x003C
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `data` - Continuation data
fn write_continue<W: Write>(writer: &mut W, data: &[u8]) -> Result<()> {
    let len = crate::utils::truncate_usize_to_u16(data.len().min(8224)); // Max record size
    write_record_header(writer, 0x003C, len)?;
    writer.write_all(&data[..len as usize])?;
    Ok(())
}

/// Write SST (Shared String Table) record with CONTINUE support
///
/// Record type: 0x00FC
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `strings` - Vector of strings to include in SST
///
/// # Implementation Notes
///
/// The SST record has a maximum size of 8224 bytes. If the SST exceeds this size,
/// CONTINUE records (0x003C) are used to store the remaining data.
///
/// This implementation properly handles string splitting across CONTINUE boundaries,
/// based on Apache POI's `SSTSerializer`.
///
/// # Errors
///
/// Returns [`crate::Error::StringTooLong`] for a string longer than the
/// 0xFFFF UTF-16 code units an SST entry can hold, before writing anything.
pub(crate) fn write_sst<W: Write, S: AsRef<str>>(
    writer: &mut W,
    strings: &[S],
    cst_total: u32,
) -> Result<()> {
    const MAX_RECORD_DATA: usize = 8224; // max data payload per record

    // Every string fits, or nothing is written.
    for s in strings {
        ensure_utf16_len_within(s.as_ref(), SHARED_STRING_UNITS, "shared string")?;
    }

    // We'll build each record's payload into a local buffer, then flush with the header
    let mut first_record = true;
    let mut buffer: Vec<u8> = Vec::with_capacity(MAX_RECORD_DATA);

    // Helper to flush current buffer as either SST or CONTINUE
    let flush = |writer: &mut W, buf: &mut Vec<u8>, first: bool| -> Result<()> {
        if buf.is_empty() {
            return Ok(());
        }
        if first {
            write_record_header(
                writer,
                0x00FC,
                crate::utils::truncate_usize_to_u16(buf.len()),
            )?;
            writer.write_all(buf)?;
        } else {
            // Delegate CONTINUE records to the helper so we keep the
            // record-writing logic in one place.
            write_continue(writer, buf)?;
        }
        buf.clear();
        Ok(())
    };

    // Initialize first SST record with cstTotal and cstUnique
    buffer.extend_from_slice(&cst_total.to_le_bytes());
    buffer.extend_from_slice(&crate::utils::truncate_usize_to_u32(strings.len()).to_le_bytes());

    // Available payload for this record
    let mut available = MAX_RECORD_DATA - buffer.len();

    // UTF-16LE bytes of the current non-ASCII string, reused across strings.
    let mut data16: Vec<u8> = Vec::new();

    for s in strings {
        let s = s.as_ref();
        let is_ascii = s.is_ascii();
        let cch: usize;
        let mut data8: &[u8] = &[];
        let high_byte_flag: u8;

        if is_ascii {
            data8 = s.as_bytes();
            cch = data8.len();
            high_byte_flag = 0x00;
        } else {
            // One pass. A UTF-8 byte never yields more than one unit, and the
            // string was checked to have at most 0xFFFF of them.
            data16.clear();
            data16.reserve(s.len().min(SHARED_STRING_UNITS) * 2);
            for unit in s.encode_utf16() {
                data16.extend_from_slice(&unit.to_le_bytes());
            }
            cch = data16.len() / 2;
            high_byte_flag = 0x01;
        }
        let cch_field = u16_len(cch, "shared string")?;

        // String header is 3 bytes (cch u16 + flags u8). Ensure it fits fully in current record.
        if available < 3 {
            // Flush current record
            flush(writer, &mut buffer, first_record)?;
            first_record = false;
            // Start CONTINUE record; for a new string we do not need leading high-byte flag
            available = MAX_RECORD_DATA;
        }

        // Write header
        buffer.extend_from_slice(&cch_field.to_le_bytes());
        buffer.push(high_byte_flag);
        available -= 3;

        // Now write character data, possibly across CONTINUE records
        if high_byte_flag == 0x00 {
            // Compressed 8-bit
            let mut offset = 0;
            while offset < cch {
                let can_write = available.min(cch - offset);
                if can_write == 0 {
                    // Flush and start CONTINUE; for continued strings, first byte is compression flag
                    flush(writer, &mut buffer, first_record)?;
                    first_record = false;
                    buffer.push(high_byte_flag); // continuation header for string
                    available = MAX_RECORD_DATA - 1;
                    continue;
                }
                buffer.extend_from_slice(&data8[offset..offset + can_write]);
                offset += can_write;
                available -= can_write;
            }
        } else {
            // UTF-16LE (2 bytes per char), do not split a character
            let total_bytes = cch * 2;
            let mut written = 0;
            while written < total_bytes {
                // space available in bytes, but keep even number to not split a char
                let mut can_write = available.min(total_bytes - written);
                if can_write == 0 {
                    // Flush and start CONTINUE; for continued strings, first byte is compression flag
                    flush(writer, &mut buffer, first_record)?;
                    first_record = false;
                    buffer.push(high_byte_flag); // continuation header for string
                    available = MAX_RECORD_DATA - 1;
                    continue;
                }
                // ensure even number of bytes
                if !can_write.is_multiple_of(2) {
                    if can_write == 1 {
                        // no space for a full char
                        flush(writer, &mut buffer, first_record)?;
                        first_record = false;
                        buffer.push(high_byte_flag);
                        available = MAX_RECORD_DATA - 1;
                        continue;
                    }
                    can_write -= 1;
                }
                buffer.extend_from_slice(&data16[written..written + can_write]);
                written += can_write;
                available -= can_write;
            }
        }
    }

    // Flush any remaining data
    flush(writer, &mut buffer, first_record)?;

    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    /// The SST writer as it was before change 0753, kept verbatim as the
    /// reference: non-ASCII strings were collected into a `Vec<u16>` and then
    /// copied into a fresh byte vector per string.
    fn legacy_write_sst<W: Write>(
        writer: &mut W,
        strings: &[String],
        cst_total: u32,
    ) -> Result<()> {
        const MAX_RECORD_DATA: usize = 8224;
        let mut first_record = true;
        let mut buffer: Vec<u8> = Vec::with_capacity(MAX_RECORD_DATA);
        let flush = |writer: &mut W, buf: &mut Vec<u8>, first: bool| -> Result<()> {
            if buf.is_empty() {
                return Ok(());
            }
            if first {
                write_record_header(
                    writer,
                    0x00FC,
                    crate::utils::truncate_usize_to_u16(buf.len()),
                )?;
                writer.write_all(buf)?;
            } else {
                write_continue(writer, buf)?;
            }
            buf.clear();
            Ok(())
        };
        buffer.extend_from_slice(&cst_total.to_le_bytes());
        buffer.extend_from_slice(&crate::utils::truncate_usize_to_u32(strings.len()).to_le_bytes());
        let mut available = MAX_RECORD_DATA - buffer.len();
        for s in strings {
            let is_ascii = s.is_ascii();
            let cch: usize;
            let mut data8: &[u8] = &[];
            let mut data16: Vec<u8> = Vec::new();
            let high_byte_flag: u8;
            if is_ascii {
                let bytes = s.as_bytes();
                cch = bytes.len().min(0xFFFF);
                data8 = &bytes[..cch];
                high_byte_flag = 0x00;
            } else {
                let utf16: Vec<u16> = s.encode_utf16().collect();
                cch = utf16.len().min(0xFFFF);
                data16.reserve_exact(cch * 2);
                for ch in utf16.iter().take(cch) {
                    data16.extend_from_slice(&ch.to_le_bytes());
                }
                high_byte_flag = 0x01;
            }
            if available < 3 {
                flush(writer, &mut buffer, first_record)?;
                first_record = false;
                available = MAX_RECORD_DATA;
            }
            buffer.extend_from_slice(&crate::utils::truncate_usize_to_u16(cch).to_le_bytes());
            buffer.push(high_byte_flag);
            available -= 3;
            if high_byte_flag == 0x00 {
                let mut offset = 0;
                while offset < cch {
                    let can_write = available.min(cch - offset);
                    if can_write == 0 {
                        flush(writer, &mut buffer, first_record)?;
                        first_record = false;
                        buffer.push(high_byte_flag);
                        available = MAX_RECORD_DATA - 1;
                        continue;
                    }
                    buffer.extend_from_slice(&data8[offset..offset + can_write]);
                    offset += can_write;
                    available -= can_write;
                }
            } else {
                let total_bytes = cch * 2;
                let mut written = 0;
                while written < total_bytes {
                    let mut can_write = available.min(total_bytes - written);
                    if can_write == 0 {
                        flush(writer, &mut buffer, first_record)?;
                        first_record = false;
                        buffer.push(high_byte_flag);
                        available = MAX_RECORD_DATA - 1;
                        continue;
                    }
                    if !can_write.is_multiple_of(2) {
                        if can_write == 1 {
                            flush(writer, &mut buffer, first_record)?;
                            first_record = false;
                            buffer.push(high_byte_flag);
                            available = MAX_RECORD_DATA - 1;
                            continue;
                        }
                        can_write -= 1;
                    }
                    buffer.extend_from_slice(&data16[written..written + can_write]);
                    written += can_write;
                    available -= can_write;
                }
            }
        }
        flush(writer, &mut buffer, first_record)?;
        Ok(())
    }

    fn assert_same_bytes(strings: &[String]) {
        let mut expected = Vec::new();
        legacy_write_sst(&mut expected, strings, 99).unwrap();
        let mut actual = Vec::new();
        write_sst(&mut actual, strings, 99).unwrap();
        assert_eq!(actual, expected);
        let borrowed: Vec<&str> = strings.iter().map(String::as_str).collect();
        let mut from_borrowed = Vec::new();
        write_sst(&mut from_borrowed, &borrowed, 99).unwrap();
        assert_eq!(from_borrowed, expected);
    }

    #[test]
    fn single_pass_encoding_writes_the_legacy_bytes() {
        let odd = "x".repeat(8_215);
        assert_same_bytes(&[]);
        assert_same_bytes(&[String::new()]);
        assert_same_bytes(&["ascii".to_string(), "àé".to_string(), "漢字".to_string()]);
        // Supplementary-plane scalars are surrogate pairs; a CONTINUE boundary
        // must fall between whole code units.
        assert_same_bytes(&[odd.clone(), "😀".repeat(5_000)]);
        assert_same_bytes(&[odd.clone() + "a", "é".repeat(9_000), "tail".to_string()]);
        assert_same_bytes(&["a".repeat(9_000), "漢".repeat(5_000)]);
        // Exactly 0xFFFF code units, the most an entry holds, in each width;
        // the last one ends with a surrogate pair.
        assert_same_bytes(&["b".repeat(0xFFFF), "é".repeat(0xFFFF)]);
        assert_same_bytes(&[format!("{}😀", "a".repeat(0xFFFD))]);
        let many: Vec<String> = (0..3_000)
            .map(|index| match index % 4 {
                0 => format!("ascii {index}"),
                1 => format!("ünïcode {index}"),
                2 => format!("😀 {index}"),
                _ => String::new(),
            })
            .collect();
        assert_same_bytes(&many);
    }

    #[test]
    fn a_string_one_unit_past_an_entry_is_refused_before_anything_is_written() {
        let fits = "fits".to_string();
        for too_long in [
            "b".repeat(0x1_0000),
            "é".repeat(0x1_0000),
            // 0xFFFE units and a surrogate pair: the old encoder kept the
            // pair's high half alone at the limit.
            format!("{}😀", "a".repeat(0xFFFE)),
        ] {
            let mut output = Vec::new();
            let result = write_sst(&mut output, &[fits.clone(), too_long], 2);
            assert!(
                matches!(
                    result,
                    Err(crate::Error::StringTooLong {
                        field: "shared string",
                        utf16_units: 0x1_0000,
                        limit: 0xFFFF,
                    })
                ),
                "{result:?}"
            );
            assert!(output.is_empty());
        }
    }

    #[test]
    fn strings_up_to_an_entry_are_written_whole() {
        for units in [0xFFFE, 0xFFFF] {
            for value in ["b".repeat(units), "漢".repeat(units)] {
                let mut output = Vec::new();
                write_sst(&mut output, std::slice::from_ref(&value), 1).unwrap();
                // cch follows the record header, cstTotal and cstUnique.
                let cch = u16::from_le_bytes([output[12], output[13]]);
                assert_eq!(usize::from(cch), units);
            }
        }
    }
}
