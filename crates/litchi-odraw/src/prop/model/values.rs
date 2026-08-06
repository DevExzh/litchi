use crate::{Error, Result};

/// A lossless OfficeArt color reference.
///
/// Indirect palette, scheme, and system colors retain their exact bit pattern.
/// Call [`Self::rgb`] only when direct RGB semantics are required.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ColorRef(u32);

impl ColorRef {
    const FLAGS_MASK: u32 = 0xFF00_0000;

    /// Wraps an exact `OfficeArtCOLORREF` value.
    pub const fn from_raw(raw: u32) -> Self {
        Self(raw)
    }

    /// Returns the exact value read from the wire.
    pub const fn raw(self) -> u32 {
        self.0
    }

    /// Returns the high-byte flags without interpreting producer extensions.
    pub const fn flags(self) -> u8 {
        (self.0 >> 24) as u8
    }

    /// Returns an RGB triple only for a direct, unflagged color.
    ///
    /// `OfficeArtCOLORREF` stores red in the low byte, followed by green and
    /// blue. Flagged values are palette, scheme, or system references and are
    /// deliberately not flattened here.
    pub const fn rgb(self) -> Option<(u8, u8, u8)> {
        if self.0 & Self::FLAGS_MASK != 0 {
            return None;
        }
        Some((
            (self.0 & 0xFF) as u8,
            ((self.0 >> 8) & 0xFF) as u8,
            ((self.0 >> 16) & 0xFF) as u8,
        ))
    }
}

/// A decoded property value that continues to borrow complex bytes.
#[derive(Debug)]
pub enum Value<'data> {
    /// A four-byte scalar, retaining its signed wire representation.
    Simple(i32),
    /// Property-specific complex bytes.
    Complex(&'data [u8]),
    /// A validated `IMsoArray`.
    Array(Array<'data>),
}

/// A validated, zero-copy `IMsoArray` view.
#[derive(Debug, Clone, Copy)]
pub struct Array<'data> {
    data: &'data [u8],
}

impl<'data> Array<'data> {
    /// Validates an entire `IMsoArray`, including its exact payload extent.
    pub fn new(data: &'data [u8]) -> Result<Self> {
        let header = data.get(..6).ok_or(Error::MalformedProperties {
            reason: "array property is shorter than its six-byte header",
        })?;
        let count = u16::from_le_bytes([header[0], header[1]]);
        let allocated = u16::from_le_bytes([header[2], header[3]]);
        if allocated < count {
            return Err(Error::MalformedProperties {
                reason: "array allocation count is smaller than its element count",
            });
        }

        let raw_size = u16::from_le_bytes([header[4], header[5]]);
        let size = if raw_size == 0xFFF0 {
            4
        } else {
            usize::from(raw_size)
        };
        let payload_len =
            usize::from(count)
                .checked_mul(size)
                .ok_or(Error::ArithmeticOverflow {
                    context: "array-property payload length",
                })?;
        let expected = 6usize
            .checked_add(payload_len)
            .ok_or(Error::ArithmeticOverflow {
                context: "array-property extent",
            })?;
        if data.len() != expected {
            return Err(Error::MalformedProperties {
                reason: "array property does not have its exact declared extent",
            });
        }
        Ok(Self { data })
    }

    /// Returns the number of encoded elements.
    #[inline]
    pub fn element_count(&self) -> u16 {
        u16::from_le_bytes([self.data[0], self.data[1]])
    }

    /// Returns the maximum number of elements the producer allocated.
    #[inline]
    pub fn element_count_in_memory(&self) -> u16 {
        u16::from_le_bytes([self.data[2], self.data[3]])
    }

    /// Returns the exact unsigned `cbElem` field.
    #[inline]
    pub fn raw_element_size(&self) -> u16 {
        u16::from_le_bytes([self.data[4], self.data[5]])
    }

    /// Returns the number of encoded bytes per element.
    #[inline]
    pub fn element_size(&self) -> usize {
        match self.raw_element_size() {
            0xFFF0 => 4,
            size => usize::from(size),
        }
    }

    /// Borrows one encoded element.
    #[inline]
    pub fn get_element(&self, index: usize) -> Option<&'data [u8]> {
        if index >= usize::from(self.element_count()) {
            return None;
        }
        let size = self.element_size();
        let start = index.checked_mul(size)?.checked_add(6)?;
        let end = start.checked_add(size)?;
        self.data.get(start..end)
    }

    /// Iterates over every encoded element in order.
    pub fn elements(&self) -> impl Iterator<Item = &'data [u8]> {
        let count = usize::from(self.element_count());
        let array = *self;
        let mut index = 0;
        std::iter::from_fn(move || {
            if index == count {
                return None;
            }
            let element = array.get_element(index);
            index += 1;
            element
        })
    }

    /// Returns the complete encoded array, including its header.
    #[inline]
    pub fn raw_data(&self) -> &'data [u8] {
        self.data
    }
}
