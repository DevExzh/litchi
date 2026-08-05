use super::model::{
    ButtonFlags, ControlFlags, ControlHeader, Error, Flags, GeneralFlags, Header, Restrictions,
    SpecificFlags, WString,
};

const TOOLBAR_SIGNATURE: u8 = 0x02;
const CONTROL_SIGNATURE: u8 = 0x03;
const VERSION: u8 = 0x01;

impl<'a> WString<'a> {
    /// Parse one `WString` and return the number of bytes consumed.
    pub fn parse_prefix(data: &'a [u8]) -> Result<(Self, usize), Error> {
        let count = *data.first().ok_or(Error::Truncated("WString length"))?;
        let payload_len = usize::from(count)
            .checked_mul(2)
            .ok_or_else(|| Error::invalid("WString length overflows usize"))?;
        let end = 1usize
            .checked_add(payload_len)
            .ok_or_else(|| Error::invalid("WString length overflows usize"))?;
        if data.len() < end {
            return Err(Error::Truncated("WString data"));
        }
        let value = Self::from_wire(&data[1..end])?;
        Ok((value, end))
    }

    /// Parse exactly one `WString`.
    pub fn parse(data: &'a [u8]) -> Result<Self, Error> {
        let (value, consumed) = Self::parse_prefix(data)?;
        if consumed != data.len() {
            return Err(Error::invalid("WString has trailing bytes"));
        }
        Ok(value)
    }

    /// Serialize the length byte and UTF-16LE payload deterministically.
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut output = Vec::with_capacity(1 + self.encoded_len());
        output.push(self.len() as u8);
        output.extend_from_slice(self.encoded_bytes());
        output
    }
}

impl Restrictions {
    /// Parse one four-byte `TBTRFlags` value.
    pub fn parse(data: &[u8]) -> Result<Self, Error> {
        if data.len() != 4 {
            return Err(if data.len() < 4 {
                Error::Truncated("TBTRFlags")
            } else {
                Error::invalid("TBTRFlags has trailing bytes")
            });
        }
        Ok(Self::from_raw(u32::from_le_bytes([
            data[0], data[1], data[2], data[3],
        ])))
    }

    /// Serialize the exact four-byte `TBTRFlags` value.
    pub fn to_bytes(self) -> [u8; 4] {
        self.raw().to_le_bytes()
    }
}

impl Flags {
    /// Parse one two-byte `TBFlags` value.
    pub fn parse(data: &[u8]) -> Result<Self, Error> {
        if data.len() != 2 {
            return Err(if data.len() < 2 {
                Error::Truncated("TBFlags")
            } else {
                Error::invalid("TBFlags has trailing bytes")
            });
        }
        Ok(Self::from_raw(u16::from_le_bytes([data[0], data[1]])))
    }

    /// Serialize the exact two-byte `TBFlags` value.
    pub fn to_bytes(self) -> [u8; 2] {
        self.raw().to_le_bytes()
    }
}

impl ControlFlags {
    /// Parse one `TBCFlags` byte.
    pub fn parse(data: &[u8]) -> Result<Self, Error> {
        if data.len() != 1 {
            return Err(if data.is_empty() {
                Error::Truncated("TBCFlags")
            } else {
                Error::invalid("TBCFlags has trailing bytes")
            });
        }
        Ok(Self::from_raw(data[0]))
    }

    /// Serialize the exact `TBCFlags` byte.
    pub const fn to_bytes(self) -> [u8; 1] {
        [self.raw()]
    }
}

impl SpecificFlags {
    /// Parse one four-byte `TBCSFlags` value.
    pub fn parse(data: &[u8]) -> Result<Self, Error> {
        if data.len() != 4 {
            return Err(if data.len() < 4 {
                Error::Truncated("TBCSFlags")
            } else {
                Error::invalid("TBCSFlags has trailing bytes")
            });
        }
        Ok(Self::from_raw(u32::from_le_bytes([
            data[0], data[1], data[2], data[3],
        ])))
    }

    /// Serialize the exact four-byte `TBCSFlags` value.
    pub fn to_bytes(self) -> [u8; 4] {
        self.raw().to_le_bytes()
    }
}

impl GeneralFlags {
    /// Parse one `TBCGIFlags` byte.
    pub fn parse(data: &[u8]) -> Result<Self, Error> {
        if data.len() != 1 {
            return Err(if data.is_empty() {
                Error::Truncated("TBCGIFlags")
            } else {
                Error::invalid("TBCGIFlags has trailing bytes")
            });
        }
        Ok(Self::from_raw(data[0]))
    }

    /// Serialize the exact `TBCGIFlags` byte.
    pub const fn to_bytes(self) -> [u8; 1] {
        [self.raw()]
    }
}

impl ButtonFlags {
    /// Parse one `TBCBSFlags` byte.
    pub fn parse(data: &[u8]) -> Result<Self, Error> {
        if data.len() != 1 {
            return Err(if data.is_empty() {
                Error::Truncated("TBCBSFlags")
            } else {
                Error::invalid("TBCBSFlags has trailing bytes")
            });
        }
        Ok(Self::from_raw(data[0]))
    }

    /// Serialize the exact `TBCBSFlags` byte.
    pub const fn to_bytes(self) -> [u8; 1] {
        [self.raw()]
    }
}

impl<'a> Header<'a> {
    /// Parse one `TB` structure and return the number of bytes consumed.
    pub fn parse_prefix(data: &'a [u8]) -> Result<(Self, usize), Error> {
        let mut cursor = Cursor::new(data);
        require_byte(
            cursor.u8("TB signature")?,
            TOOLBAR_SIGNATURE,
            "TB signature",
        )?;
        require_byte(cursor.u8("TB version")?, VERSION, "TB version")?;
        let control_count = cursor.i16("TB cCL")?;
        if cursor.u32("TB toolbar id")? != 1 {
            return Err(Error::invalid("TB toolbar id must be 1"));
        }
        let restrictions = Restrictions::from_raw(cursor.u32("TBTRFlags")?);
        let rows_default = cursor.u16("TB cRowsDefault")?;
        if rows_default > u16::from(u8::MAX) {
            return Err(Error::invalid("TB cRowsDefault exceeds 255"));
        }
        let flags = Flags::from_raw(cursor.u16("TBFlags")?);
        let (name, consumed) = WString::parse_prefix(cursor.remaining())?;
        cursor.advance(consumed)?;
        Ok((
            Self::from_decoded(control_count, restrictions, rows_default, flags, name),
            cursor.position(),
        ))
    }

    /// Parse exactly one `TB` structure.
    pub fn parse(data: &'a [u8]) -> Result<Self, Error> {
        let (value, consumed) = Self::parse_prefix(data)?;
        if consumed != data.len() {
            return Err(Error::invalid("TB has trailing bytes"));
        }
        Ok(value)
    }

    /// Serialize the `TB` structure deterministically.
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut output = Vec::with_capacity(17 + self.name().encoded_len());
        output.push(TOOLBAR_SIGNATURE);
        output.push(VERSION);
        output.extend_from_slice(&self.control_count().to_le_bytes());
        output.extend_from_slice(&1u32.to_le_bytes());
        output.extend_from_slice(&self.restrictions().to_bytes());
        output.extend_from_slice(&self.rows_default().to_le_bytes());
        output.extend_from_slice(&self.flags().to_bytes());
        output.extend_from_slice(&self.name().to_bytes());
        output
    }
}

impl ControlHeader {
    /// Parse one `TBCHeader` and return the number of bytes consumed.
    pub fn parse_prefix(data: &[u8]) -> Result<(Self, usize), Error> {
        let mut cursor = Cursor::new(data);
        require_byte(
            cursor.u8("TBCHeader signature")?,
            CONTROL_SIGNATURE,
            "TBCHeader signature",
        )?;
        require_byte(
            cursor.u8("TBCHeader version")?,
            VERSION,
            "TBCHeader version",
        )?;
        let flags = ControlFlags::from_raw(cursor.u8("TBCFlags")?);
        let control_type = super::model::ControlType::from_raw(cursor.u8("TBCHeader type")?);
        let control_id = cursor.u16("TBCHeader control id")?;
        let specifics = SpecificFlags::from_raw(cursor.u32("TBCSFlags")?);
        let priority = cursor.u8("TBCHeader priority")?;
        if priority > 7 {
            return Err(Error::invalid("TBCHeader priority exceeds 7"));
        }
        let dimensions = if flags.save_dimensions() {
            Some(super::model::Dimensions::new(
                cursor.u16("TBCHeader width")?,
                cursor.u16("TBCHeader height")?,
            ))
        } else {
            None
        };
        Ok((
            Self::from_decoded(
                control_type,
                control_id,
                flags,
                specifics,
                priority,
                dimensions,
            ),
            cursor.position(),
        ))
    }

    /// Parse exactly one `TBCHeader`.
    pub fn parse(data: &[u8]) -> Result<Self, Error> {
        let (value, consumed) = Self::parse_prefix(data)?;
        if consumed != data.len() {
            return Err(Error::invalid("TBCHeader has trailing bytes"));
        }
        Ok(value)
    }

    /// Serialize the `TBCHeader` structure deterministically.
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut output = Vec::with_capacity(if self.dimensions().is_some() { 15 } else { 11 });
        output.push(CONTROL_SIGNATURE);
        output.push(VERSION);
        output.extend_from_slice(&self.flags().to_bytes());
        output.push(self.control_type().raw());
        output.extend_from_slice(&self.control_id().to_le_bytes());
        output.extend_from_slice(&self.specifics().to_bytes());
        output.push(self.priority());
        if let Some(dimensions) = self.dimensions() {
            output.extend_from_slice(&dimensions.width().to_le_bytes());
            output.extend_from_slice(&dimensions.height().to_le_bytes());
        }
        output
    }
}

fn require_byte(actual: u8, expected: u8, field: &'static str) -> Result<(), Error> {
    if actual == expected {
        Ok(())
    } else {
        Err(Error::invalid(format!(
            "{field} must be 0x{expected:02X}, got 0x{actual:02X}"
        )))
    }
}

struct Cursor<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> Cursor<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn remaining(&self) -> &'a [u8] {
        &self.data[self.offset..]
    }

    fn position(&self) -> usize {
        self.offset
    }

    fn advance(&mut self, count: usize) -> Result<(), Error> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| Error::invalid("toolbar cursor offset overflows usize"))?;
        if end > self.data.len() {
            return Err(Error::Truncated("toolbar structure"));
        }
        self.offset = end;
        Ok(())
    }

    fn take(&mut self, count: usize, field: &'static str) -> Result<&'a [u8], Error> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| Error::invalid(format!("{field} length overflows usize")))?;
        if end > self.data.len() {
            return Err(Error::Truncated(field));
        }
        let value = &self.data[self.offset..end];
        self.offset = end;
        Ok(value)
    }

    fn u8(&mut self, field: &'static str) -> Result<u8, Error> {
        Ok(self.take(1, field)?[0])
    }

    fn i16(&mut self, field: &'static str) -> Result<i16, Error> {
        let value = self.take(2, field)?;
        Ok(i16::from_le_bytes([value[0], value[1]]))
    }

    fn u16(&mut self, field: &'static str) -> Result<u16, Error> {
        let value = self.take(2, field)?;
        Ok(u16::from_le_bytes([value[0], value[1]]))
    }

    fn u32(&mut self, field: &'static str) -> Result<u32, Error> {
        let value = self.take(4, field)?;
        Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
    }
}
