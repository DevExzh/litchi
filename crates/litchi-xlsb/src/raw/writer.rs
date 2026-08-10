//! Validated BIFF12 record writer.

use std::io::Write;

use super::cursor::decode_rk;
use super::{Error, Kind, Limits, Result};

/// Forward-only BIFF12 writer.
pub struct Writer<W: Write> {
    writer: W,
    limits: Limits,
}

impl<W: Write> Writer<W> {
    /// Write with safe default limits.
    #[must_use]
    pub const fn new(writer: W) -> Self {
        Self::with_limits(writer, Limits::DEFAULT)
    }

    /// Write with explicit finite limits.
    #[must_use]
    pub const fn with_limits(writer: W, limits: Limits) -> Self {
        Self { writer, limits }
    }

    /// Write one complete record.
    pub fn write_record(&mut self, kind: Kind, payload: &[u8]) -> Result<()> {
        self.write_header(kind, payload.len())?;
        self.writer.write_all(payload)?;
        Ok(())
    }

    /// Write a validated record header.
    pub fn write_header(&mut self, kind: Kind, len: usize) -> Result<()> {
        if len > self.limits.payload() {
            return Err(Error::PayloadLimit {
                length: len,
                limit: self.limits.payload(),
                offset: 0,
            });
        }
        if len > super::record::MAX_WIRE_PAYLOAD {
            return Err(Error::LengthOverflow {
                what: "record payload",
                length: len,
            });
        }
        self.write_kind(kind)?;
        self.write_len(len)
    }

    /// Stream an `XLWideString` without constructing an intermediate UTF-16 buffer.
    pub fn write_wide_string(&mut self, value: &str) -> Result<()> {
        let units = value.encode_utf16().count();
        if units > self.limits.string_units() {
            return Err(Error::StringLimit {
                units,
                limit: self.limits.string_units(),
                offset: 0,
            });
        }
        let wire_units = u32::try_from(units).map_err(|_| Error::LengthOverflow {
            what: "UTF-16 string",
            length: units,
        })?;
        self.write_u32(wire_units)?;
        for unit in value.encode_utf16() {
            self.write_u16(unit)?;
        }
        Ok(())
    }

    /// Write one byte.
    pub fn write_u8(&mut self, value: u8) -> Result<()> {
        self.writer.write_all(&[value])?;
        Ok(())
    }

    /// Write a little-endian `u16`.
    pub fn write_u16(&mut self, value: u16) -> Result<()> {
        self.writer.write_all(&value.to_le_bytes())?;
        Ok(())
    }

    /// Write a little-endian `u32`.
    pub fn write_u32(&mut self, value: u32) -> Result<()> {
        self.writer.write_all(&value.to_le_bytes())?;
        Ok(())
    }

    /// Write a little-endian `i32`.
    pub fn write_i32(&mut self, value: i32) -> Result<()> {
        self.writer.write_all(&value.to_le_bytes())?;
        Ok(())
    }

    /// Write a little-endian `f64`.
    pub fn write_f64(&mut self, value: f64) -> Result<()> {
        self.writer.write_all(&value.to_le_bytes())?;
        Ok(())
    }

    /// Write an exactly representable RK-compressed number.
    ///
    /// Following the checked-in `[MS-XLSB]` section 2.5.123, this refuses a
    /// value when every legal RK form would discard significant bits.
    pub fn write_rk(&mut self, value: f64) -> Result<()> {
        self.write_u32(Self::f64_to_rk(value)?)
    }

    /// Flush the destination.
    pub fn flush(&mut self) -> Result<()> {
        self.writer.flush()?;
        Ok(())
    }

    /// Borrow the destination.
    #[must_use]
    pub const fn get_ref(&self) -> &W {
        &self.writer
    }

    /// Mutably borrow the destination.
    #[must_use]
    pub fn get_mut(&mut self) -> &mut W {
        &mut self.writer
    }

    /// Consume the writer and return its destination.
    #[must_use]
    pub fn finish(self) -> W {
        self.writer
    }

    fn write_kind(&mut self, kind: Kind) -> Result<()> {
        let value = kind.get();
        if value < 0x80 {
            self.write_u8(value as u8)
        } else {
            self.write_u8((value as u8 & 0x7f) | 0x80)?;
            self.write_u8((value >> 7) as u8)
        }
    }

    fn write_len(&mut self, mut value: usize) -> Result<()> {
        if value > super::record::MAX_WIRE_PAYLOAD {
            return Err(Error::LengthOverflow {
                what: "record payload",
                length: value,
            });
        }
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            self.write_u8(byte)?;
            if value == 0 {
                return Ok(());
            }
        }
    }

    fn f64_to_rk(value: f64) -> Result<u32> {
        if let Some(integer) = signed_30(value) {
            let rk = ((integer as u32) << 2) | 0x02;
            if exact_rk(value, rk) {
                return Ok(rk);
            }
        }
        let scaled = value * 100.0;
        if let Some(integer) = signed_30(scaled) {
            let rk = ((integer as u32) << 2) | 0x03;
            if exact_rk(value, rk) {
                return Ok(rk);
            }
        }

        let rk = ((value.to_bits() >> 32) as u32) & 0xffff_fffc;
        if exact_rk(value, rk) {
            return Ok(rk);
        }
        let scaled_rk = (((scaled.to_bits() >> 32) as u32) & 0xffff_fffc) | 0x01;
        if exact_rk(value, scaled_rk) {
            return Ok(scaled_rk);
        }
        Err(Error::UnrepresentableRk {
            bits: value.to_bits(),
        })
    }
}

fn exact_rk(value: f64, rk: u32) -> bool {
    decode_rk(rk).to_bits() == value.to_bits()
}

fn signed_30(value: f64) -> Option<i32> {
    const MIN: i32 = -(1 << 29);
    const MAX: i32 = (1 << 29) - 1;

    if value < f64::from(MIN) || value > f64::from(MAX) {
        return None;
    }
    let integer = value as i32;
    (f64::from(integer).to_bits() == value.to_bits()).then_some(integer)
}
