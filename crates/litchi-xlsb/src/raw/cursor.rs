#![allow(
    clippy::cast_possible_wrap,
    clippy::map_err_ignore,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, normalization into the module's stable typed public error to this codec boundary"
)]

//! Bounds-checked cursor over one record payload.

use std::char;

use super::{Error, Limits, Result, Stage};

/// A borrowing cursor for validated scalar and string reads.
pub struct Cursor<'a> {
    data: &'a [u8],
    offset: usize,
    context: &'static str,
    limits: Limits,
}

impl<'a> Cursor<'a> {
    /// Read with safe default limits.
    #[must_use]
    pub const fn new(data: &'a [u8], context: &'static str) -> Self {
        Self::with_limits(data, context, Limits::DEFAULT)
    }

    /// Read with explicit finite limits.
    #[must_use]
    pub const fn with_limits(data: &'a [u8], context: &'static str, limits: Limits) -> Self {
        Self {
            data,
            offset: 0,
            context,
            limits,
        }
    }

    /// Construct a cursor after validating its raw limit profile.
    pub fn try_with_limits(
        data: &'a [u8],
        context: &'static str,
        limits: Limits,
    ) -> Result<Self> {
        match limits.validate() {
            Ok(limits) => Ok(Self::with_limits(data, context, limits)),
            Err(error) => Err(error),
        }
    }

    /// Diagnostic context supplied by the caller.
    #[must_use]
    pub const fn context(&self) -> &'static str {
        self.context
    }

    /// Current cursor offset.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.offset
    }

    /// Number of unread bytes.
    #[must_use]
    pub fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.offset)
    }

    /// Require that `needed` bytes remain.
    pub fn guard(&self, needed: usize) -> Result<()> {
        self.limits.validate()?;
        let end = self
            .offset
            .checked_add(needed)
            .ok_or(Error::LengthOverflow {
                what: "payload cursor",
                length: needed,
            })?;
        if end > self.data.len() {
            return Err(Error::Truncated {
                stage: Stage::Value,
                offset: self.offset,
                needed,
                available: self.remaining(),
            });
        }
        Ok(())
    }

    /// Advance without copying bytes.
    pub fn skip(&mut self, len: usize) -> Result<()> {
        self.guard(len)?;
        self.offset = self.offset.checked_add(len).ok_or(Error::LengthOverflow {
            what: "payload cursor",
            length: len,
        })?;
        Ok(())
    }

    /// Lend the next `len` bytes.
    pub fn read_bytes(&mut self, len: usize) -> Result<&'a [u8]> {
        self.guard(len)?;
        let end = self.offset.checked_add(len).ok_or(Error::LengthOverflow {
            what: "payload cursor",
            length: len,
        })?;
        let bytes = self.data.get(self.offset..end).ok_or(Error::Truncated {
            stage: Stage::Value,
            offset: self.offset,
            needed: len,
            available: self.remaining(),
        })?;
        self.offset = end;
        Ok(bytes)
    }

    /// Read one byte.
    pub fn read_u8(&mut self) -> Result<u8> {
        let offset = self.offset;
        self.skip(1)?;
        self.data.get(offset).copied().ok_or(Error::Truncated {
            stage: Stage::Value,
            offset,
            needed: 1,
            available: 0,
        })
    }

    /// Read a little-endian `u16`.
    pub fn read_u16(&mut self) -> Result<u16> {
        Ok(u16::from_le_bytes(self.array()?))
    }

    /// Read a little-endian `i16`.
    pub fn read_i16(&mut self) -> Result<i16> {
        Ok(i16::from_le_bytes(self.array()?))
    }

    /// Read a little-endian `u32`.
    pub fn read_u32(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes(self.array()?))
    }

    /// Read a little-endian `i32`.
    pub fn read_i32(&mut self) -> Result<i32> {
        Ok(i32::from_le_bytes(self.array()?))
    }

    /// Read a little-endian `f64`.
    pub fn read_f64(&mut self) -> Result<f64> {
        Ok(f64::from_le_bytes(self.array()?))
    }

    /// Read a byte Boolean, rejecting values other than zero and one.
    pub fn read_bool8(&mut self) -> Result<bool> {
        let offset = self.offset;
        let value = self.read_u8()?;
        match value {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(Error::InvalidBool {
                value: u32::from(value),
                offset,
            }),
        }
    }

    /// Read a 32-bit Boolean, rejecting values other than zero and one.
    pub fn read_bool32(&mut self) -> Result<bool> {
        let offset = self.offset;
        let value = self.read_u32()?;
        match value {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(Error::InvalidBool { value, offset }),
        }
    }

    /// Decode one RK-compressed number per `[MS-XLSB]` section 2.5.123.
    pub fn read_rk(&mut self) -> Result<f64> {
        self.read_u32().map(decode_rk)
    }

    /// Strictly decode one `XLWideString`.
    pub fn read_wide_string(&mut self) -> Result<String> {
        let length_offset = self.offset;
        let units_u32 = self.read_u32()?;
        let units = usize::try_from(units_u32).map_err(|_| Error::LengthOverflow {
            what: "UTF-16 string",
            length: usize::MAX,
        })?;
        if units > self.limits.string_units() {
            return Err(Error::StringLimit {
                units,
                limit: self.limits.string_units(),
                offset: length_offset,
            });
        }
        let byte_len = units.checked_mul(2).ok_or(Error::LengthOverflow {
            what: "UTF-16 string",
            length: units,
        })?;
        let string_offset = self.offset;
        let bytes = self.read_bytes(byte_len)?;
        let mut output = String::with_capacity(units);
        for decoded in char::decode_utf16(Utf16Le::new(bytes)) {
            let value = decoded.map_err(|_| Error::InvalidUtf16 {
                offset: string_offset,
            })?;
            output.push(value);
        }
        Ok(output)
    }

    /// Strictly decode one `XLNullableWideString`.
    pub fn read_nullable_wide_string(&mut self) -> Result<Option<String>> {
        self.guard(4)?;
        let prefix = self
            .data
            .get(self.offset..self.offset.saturating_add(4))
            .ok_or(Error::Truncated {
                stage: Stage::Value,
                offset: self.offset,
                needed: 4,
                available: self.remaining(),
            })?;
        let marker = u32::from_le_bytes(prefix.try_into().map_err(|_| Error::Truncated {
            stage: Stage::Value,
            offset: self.offset,
            needed: 4,
            available: prefix.len(),
        })?);
        if marker == u32::MAX {
            self.skip(4)?;
            return Ok(None);
        }
        self.read_wide_string().map(Some)
    }

    /// Lend one length-prefixed byte blob without copying.
    pub fn read_blob(&mut self) -> Result<&'a [u8]> {
        let len_u32 = self.read_u32()?;
        let len = usize::try_from(len_u32).map_err(|_| Error::LengthOverflow {
            what: "byte blob",
            length: usize::MAX,
        })?;
        if len > self.limits.payload() {
            return Err(Error::PayloadLimit {
                length: len,
                limit: self.limits.payload(),
                offset: self.offset.saturating_sub(4),
            });
        }
        self.read_bytes(len)
    }

    /// Require that the complete payload was consumed.
    pub fn finish(&self) -> Result<()> {
        self.limits.validate()?;
        let remaining = self.remaining();
        if remaining != 0 {
            return Err(Error::Trailing {
                context: self.context,
                offset: self.offset,
                remaining,
            });
        }
        Ok(())
    }

    fn array<const N: usize>(&mut self) -> Result<[u8; N]> {
        let offset = self.offset;
        let bytes = self.read_bytes(N)?;
        bytes.try_into().map_err(|_| Error::Truncated {
            stage: Stage::Value,
            offset,
            needed: N,
            available: bytes.len(),
        })
    }
}

pub(super) fn decode_rk(rk: u32) -> f64 {
    let mut value = if rk & 0x02 != 0 {
        f64::from((rk as i32) >> 2)
    } else {
        f64::from_bits(u64::from(rk & 0xffff_fffc) << 32)
    };
    if rk & 0x01 != 0 {
        value /= 100.0;
    }
    value
}

struct Utf16Le<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Utf16Le<'a> {
    const fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }
}

impl Iterator for Utf16Le<'_> {
    type Item = u16;

    fn next(&mut self) -> Option<Self::Item> {
        let end = self.offset.checked_add(2)?;
        let bytes = self.bytes.get(self.offset..end)?;
        let low = bytes.first().copied()?;
        let high = bytes.get(1).copied()?;
        self.offset = end;
        Some(u16::from_le_bytes([low, high]))
    }
}
