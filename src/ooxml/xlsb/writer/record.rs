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
    fn test_write_variable_length_u16() {
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
    fn test_write_wide_string() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        writer.write_wide_string("A").unwrap();
        // Length (1) + UTF-16LE 'A' (0x41, 0x00)
        assert_eq!(buffer, vec![0x01, 0x00, 0x00, 0x00, 0x41, 0x00]);
    }

    #[test]
    fn test_f64_to_rk() {
        // Integer value
        let rk = RecordWriter::<Vec<u8>>::f64_to_rk(100.0);
        assert_eq!(rk & 0x03, 0x01); // Integer flag

        // Integer divisible by 100
        let rk = RecordWriter::<Vec<u8>>::f64_to_rk(10000.0);
        assert_eq!(rk & 0x03, 0x03); // Integer / 100 flag
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
}
