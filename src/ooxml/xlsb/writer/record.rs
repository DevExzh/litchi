//! Binary record writer for XLSB format
//!
//! Implements variable-length encoding for record types and sizes
//! according to the MS-XLSB specification.

use crate::ooxml::xlsb::error::XlsbResult;
use std::io::Write;

/// XLSB record writer with variable-length encoding support
pub struct RecordWriter<W: Write> {
    writer: W,
}

impl<W: Write> RecordWriter<W> {
    /// Create a new record writer
    pub fn new(writer: W) -> Self {
        RecordWriter { writer }
    }

    /// Write a complete record with header and data
    ///
    /// # Arguments
    ///
    /// * `record_type` - The record type identifier
    /// * `data` - The record data bytes
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsb::writer::RecordWriter;
    /// use std::io::Cursor;
    ///
    /// let mut buffer = Vec::new();
    /// let mut writer = RecordWriter::new(&mut buffer);
    /// writer.write_record(0x0001, &[0x00, 0x00, 0x00, 0x00])?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn write_record(&mut self, record_type: u16, data: &[u8]) -> XlsbResult<()> {
        self.write_record_header(record_type, data.len())?;
        self.writer.write_all(data)?;
        Ok(())
    }

    /// Write record header with variable-length encoding
    fn write_record_header(&mut self, record_type: u16, data_len: usize) -> XlsbResult<()> {
        // Write record type (1-3 bytes, variable-length encoded)
        self.write_variable_length_u16(record_type)?;

        // Write data length (1-4 bytes, variable-length encoded)
        self.write_variable_length_usize(data_len)?;

        Ok(())
    }

    /// Write a u16 value with variable-length encoding (1-3 bytes)
    fn write_variable_length_u16(&mut self, mut value: u16) -> XlsbResult<()> {
        // First byte: lower 7 bits + continuation bit
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;

        if value > 0 {
            byte |= 0x80; // Set continuation bit
        }
        self.writer.write_all(&[byte])?;

        if value > 0 {
            // Second byte
            byte = (value & 0x7F) as u8;
            value >>= 7;

            if value > 0 {
                byte |= 0x80;
            }
            self.writer.write_all(&[byte])?;

            if value > 0 {
                // Third byte
                byte = (value & 0x7F) as u8;
                self.writer.write_all(&[byte])?;
            }
        }

        Ok(())
    }

    /// Write a usize value with variable-length encoding (1-4 bytes)
    ///
    /// XLSB uses at most 4 bytes for the length field (28 bits). Return an error
    /// if the value exceeds this range.
    fn write_variable_length_usize(&mut self, mut value: usize) -> XlsbResult<()> {
        // Reject values requiring more than 28 bits
        if value >> 28 != 0 {
            return Err(crate::ooxml::xlsb::error::XlsbError::InvalidLength {
                expected: (1 << 28) - 1,
                found: value,
            });
        }

        let mut bytes_written = 0u8;
        loop {
            let mut byte = (value & 0x7F) as u8;
            value >>= 7;

            if value > 0 {
                byte |= 0x80; // Set continuation bit
            }

            self.writer.write_all(&[byte])?;
            bytes_written += 1;

            if value == 0 {
                break;
            }

            // Safety: we already checked value <= 28 bits, so this will not exceed 4 bytes
            debug_assert!(bytes_written < 4);
        }

        Ok(())
    }

    /// Write a wide string (UTF-16LE with length prefix)
    pub fn write_wide_string(&mut self, s: &str) -> XlsbResult<()> {
        // Convert to UTF-16
        let utf16: Vec<u16> = s.encode_utf16().collect();

        // Write length (number of UTF-16 code units)
        self.write_u32(utf16.len() as u32)?;

        // Write UTF-16LE bytes
        for code_unit in utf16 {
            self.write_u16(code_unit)?;
        }

        Ok(())
    }

    /// Write a u8 value
    pub fn write_u8(&mut self, value: u8) -> XlsbResult<()> {
        self.writer.write_all(&[value])?;
        Ok(())
    }

    /// Write a u16 value (little-endian)
    pub fn write_u16(&mut self, value: u16) -> XlsbResult<()> {
        self.writer.write_all(&value.to_le_bytes())?;
        Ok(())
    }

    /// Write a u32 value (little-endian)
    pub fn write_u32(&mut self, value: u32) -> XlsbResult<()> {
        self.writer.write_all(&value.to_le_bytes())?;
        Ok(())
    }

    /// Write an i32 value (little-endian)
    pub fn write_i32(&mut self, value: i32) -> XlsbResult<()> {
        self.writer.write_all(&value.to_le_bytes())?;
        Ok(())
    }

    /// Write a f64 value (little-endian)
    pub fn write_f64(&mut self, value: f64) -> XlsbResult<()> {
        self.writer.write_all(&value.to_le_bytes())?;
        Ok(())
    }

    /// Write an RK value (compressed number format)
    ///
    /// RK values are a compressed format for numbers that can be represented
    /// as integers or with limited precision.
    pub fn write_rk(&mut self, value: f64) -> XlsbResult<()> {
        let rk = Self::f64_to_rk(value);
        self.write_u32(rk)?;
        Ok(())
    }

    /// Convert f64 to RK format
    fn f64_to_rk(value: f64) -> u32 {
        // Try to encode as integer first
        if value == value.floor() && value >= i32::MIN as f64 && value <= i32::MAX as f64 {
            let int_val = value as i32;
            // Use the /100 encoding only when the magnitude is large enough to benefit
            if int_val % 100 == 0 && int_val.abs() >= 10_000 {
                let div_val = int_val / 100;
                return ((div_val as u32) << 2) | 0x03; // Integer / 100
            }
            return ((int_val as u32) << 2) | 0x01; // Integer
        }

        // Encode as float (30-bit precision)
        let bits = value.to_bits();
        let masked = (bits >> 34) as u32;
        masked << 2 // Float format
    }

    /// Flush the underlying writer
    pub fn flush(&mut self) -> XlsbResult<()> {
        self.writer.flush()?;
        Ok(())
    }

    /// Get a reference to the inner writer
    pub fn inner(&self) -> &W {
        &self.writer
    }

    /// Get a mutable reference to the inner writer
    pub fn inner_mut(&mut self) -> &mut W {
        &mut self.writer
    }

    /// Consume the writer and return the inner writer
    pub fn into_inner(self) -> W {
        self.writer
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_new() {
        let buffer = Vec::new();
        let writer = RecordWriter::new(buffer);
        // Just verify it creates without error
        let _ = writer;
    }

    #[test]
    fn test_write_record() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        // Write a simple record
        writer
            .write_record(0x0001, &[0x00, 0x00, 0x00, 0x00])
            .unwrap();

        // Should have type (1 byte) + length (1 byte) + data (4 bytes) = 6 bytes
        assert_eq!(buffer.len(), 6);
    }

    #[test]
    fn test_write_u8() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_u8(0x42).unwrap();
        assert_eq!(buffer, vec![0x42]);
    }

    #[test]
    fn test_write_u16() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_u16(0x1234).unwrap();
        assert_eq!(buffer, vec![0x34, 0x12]); // Little-endian
    }

    #[test]
    fn test_write_u32() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_u32(0x12345678).unwrap();
        assert_eq!(buffer, vec![0x78, 0x56, 0x34, 0x12]); // Little-endian
    }

    #[test]
    fn test_write_i32() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_i32(-1).unwrap();
        assert_eq!(buffer, vec![0xFF, 0xFF, 0xFF, 0xFF]);
    }

    #[test]
    fn test_write_f64() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_f64(3.14159).unwrap();
        assert_eq!(buffer.len(), 8);
    }

    #[test]
    fn test_write_rk() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_rk(42.0).unwrap();
        assert_eq!(buffer.len(), 4);
    }

    #[test]
    fn test_flush() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_u8(0x01).unwrap();
        writer.flush().unwrap();
        // Flush should succeed
    }

    #[test]
    fn test_inner() {
        let buffer = Vec::new();
        let writer = RecordWriter::new(buffer);

        let inner = writer.inner();
        assert!(inner.is_empty());
    }

    #[test]
    fn test_inner_mut() {
        let buffer = Vec::new();
        let mut writer = RecordWriter::new(buffer);

        let inner = writer.inner_mut();
        inner.push(0x01);
        assert_eq!(inner.len(), 1);
    }

    #[test]
    fn test_into_inner() {
        let buffer = Vec::new();
        let mut writer = RecordWriter::new(buffer);

        writer.write_u8(0x42).unwrap();
        let inner = writer.into_inner();
        assert_eq!(inner, vec![0x42]);
    }

    #[test]
    fn test_write_variable_length_u16_single_byte() {
        let mut buffer = Vec::new();
        {
            let mut writer = RecordWriter::new(&mut buffer);
            // Test small value (1 byte)
            writer.write_variable_length_u16(0x0001).unwrap();
        }
        assert_eq!(buffer, vec![0x01]);

        buffer.clear();

        {
            let mut writer = RecordWriter::new(&mut buffer);
            // Test larger value (2 bytes)
            writer.write_variable_length_u16(0x0080).unwrap();
        }
        assert_eq!(buffer, vec![0x80, 0x01]);
    }

    #[test]
    fn test_write_variable_length_u16_two_bytes() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        // 0x2000 = 8192 = 0b10000000000000
        // First byte: 0b0000000 = 0x00 with continuation
        // Second byte: 0b1000000 = 0x40
        writer.write_variable_length_u16(0x2000).unwrap();
        assert_eq!(buffer.len(), 2);
    }

    #[test]
    fn test_write_variable_length_u16_three_bytes() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        // 0x4000 requires 3 bytes in 7-bit varint
        writer.write_variable_length_u16(0x4000).unwrap();
        assert_eq!(buffer.len(), 3);
    }

    #[test]
    fn test_write_variable_length_u16_max() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        // u16::MAX = 0xFFFF
        writer.write_variable_length_u16(u16::MAX).unwrap();
        assert_eq!(buffer.len(), 3);
    }

    #[test]
    fn test_write_wide_string() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_wide_string("A").unwrap();
        // Length (1) + UTF-16LE 'A' (0x41, 0x00)
        assert_eq!(buffer, vec![0x01, 0x00, 0x00, 0x00, 0x41, 0x00]);
    }

    #[test]
    fn test_write_wide_string_empty() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_wide_string("").unwrap();
        // Length (0) = 4 bytes
        assert_eq!(buffer, vec![0x00, 0x00, 0x00, 0x00]);
    }

    #[test]
    fn test_write_wide_string_unicode() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_wide_string("Hello 世界").unwrap();
        // Length (8 chars) + UTF-16LE bytes
        assert!(buffer.len() > 4);
    }

    #[test]
    fn test_f64_to_rk_integer() {
        // Integer value
        let rk = RecordWriter::<Vec<u8>>::f64_to_rk(100.0);
        assert_eq!(rk & 0x03, 0x01); // Integer flag
    }

    #[test]
    fn test_f64_to_rk_integer_div100() {
        // Integer divisible by 100
        let rk = RecordWriter::<Vec<u8>>::f64_to_rk(10000.0);
        assert_eq!(rk & 0x03, 0x03); // Integer / 100 flag
    }

    #[test]
    fn test_f64_to_rk_float() {
        // Non-integer value
        let rk = RecordWriter::<Vec<u8>>::f64_to_rk(3.14);
        assert_eq!(rk & 0x03, 0x00); // Float flag
    }

    #[test]
    fn test_f64_to_rk_negative_integer() {
        // Negative integer
        let rk = RecordWriter::<Vec<u8>>::f64_to_rk(-100.0);
        assert_eq!(rk & 0x03, 0x01); // Integer flag
    }

    #[test]
    fn test_f64_to_rk_zero() {
        let rk = RecordWriter::<Vec<u8>>::f64_to_rk(0.0);
        assert_eq!(rk & 0x03, 0x01); // Integer flag (0 is an integer)
    }

    #[test]
    fn test_header_roundtrip_large_values() {
        // Write a record with a multi-byte type and multi-byte length
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        let record_type: u16 = 0x1234; // requires multiple bytes in 7-bit varint
        let data = vec![0x55u8; 300]; // length requires multiple bytes in varint

        writer.write_record(record_type, &data).unwrap();

        // Read back header using the reader implementation
        let mut cursor = Cursor::new(&buffer);
        let header = crate::ooxml::xlsb::records::XlsbRecordHeader::read(&mut cursor).unwrap();

        assert_eq!(header.record_type, record_type);
        assert_eq!(header.data_len, data.len());
    }

    #[test]
    fn test_header_roundtrip_small_values() {
        // Write a record with single-byte type and small length
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        let record_type: u16 = 0x003F; // fits in one byte varint
        let data = [0xAAu8; 5];

        writer.write_record(record_type, &data).unwrap();

        // Read back header using the reader implementation
        let mut cursor = Cursor::new(&buffer);
        let header = crate::ooxml::xlsb::records::XlsbRecordHeader::read(&mut cursor).unwrap();

        assert_eq!(header.record_type, record_type);
        assert_eq!(header.data_len, data.len());
    }

    #[test]
    fn test_write_variable_length_usize_single_byte() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_variable_length_usize(0x7F).unwrap();
        assert_eq!(buffer.len(), 1);
    }

    #[test]
    fn test_write_variable_length_usize_multi_byte() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_variable_length_usize(0x100).unwrap();
        assert!(buffer.len() >= 2);
    }
}
