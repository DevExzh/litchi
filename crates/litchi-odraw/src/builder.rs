//! Authoring builders for typed `OfficeArt` records.
//!
//! Builders own or borrow the bytes staged for a record.  The low-level
//! framing and output functions remain in [`crate::write`]; keeping drafts in
//! this module makes the authoring namespace distinct from those primitives.

use std::{
    borrow::Cow,
    io::{self, Write},
};

use zerocopy::IntoBytes as _;

use crate::write::{self, Atom, Sp};
use crate::{
    Limits,
    prop::Id,
    shape::{Flags, Native},
};

/// Input accepted by the raw property-table builder.
pub trait PropertyKey {
    /// Returns the exact property identifier, including any wire flags.
    fn raw(self) -> u16;
}

impl PropertyKey for Id {
    fn raw(self) -> u16 {
        Id::raw(self)
    }
}

impl PropertyKey for u16 {
    fn raw(self) -> u16 {
        self
    }
}

#[derive(Debug)]
enum BuiltValue<'data> {
    Simple(i32),
    Blip(i32),
    Complex(Cow<'data, [u8]>),
}

/// Move-or-borrow builder for an `OfficeArt` Opt property table.
///
/// Owned complex values are moved into the builder and borrowed values retain
/// their input lifetime. The builder deliberately does not implement `Clone`,
/// preventing an accidental deep copy of owned complex property bytes.
#[derive(Debug, Default)]
pub struct PropertyBuilder<'data> {
    properties: Vec<(u16, BuiltValue<'data>)>,
}

impl<'data> PropertyBuilder<'data> {
    /// Creates an empty property table.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            properties: Vec::new(),
        }
    }

    /// Appends a simple property.
    pub fn add_simple<K: PropertyKey>(&mut self, id: K, value: i32) -> &mut Self {
        self.properties.push((id.raw(), BuiltValue::Simple(value)));
        self
    }

    /// Appends a BLIP-store identifier property.
    pub fn add_blip_id<K: PropertyKey>(&mut self, id: K, value: i32) -> &mut Self {
        self.properties.push((id.raw(), BuiltValue::Blip(value)));
        self
    }

    /// Appends a complex property by borrowing a slice or moving an owned vector.
    pub fn add_complex<K, D>(&mut self, id: K, data: D) -> &mut Self
    where
        K: PropertyKey,
        D: Into<Cow<'data, [u8]>>,
    {
        self.properties
            .push((id.raw(), BuiltValue::Complex(data.into())));
        self
    }

    /// Returns the exact encoded record size after checked arithmetic.
    ///
    /// # Errors
    ///
    /// Returns an error if the encoded record size overflows `usize` or if a
    /// complex property exceeds `i32::MAX` bytes.
    pub fn size(&self) -> io::Result<usize> {
        let headers = self
            .properties
            .len()
            .checked_mul(6)
            .ok_or_else(|| write::invalid_input("property-header size overflow"))?;
        let base = 8_usize
            .checked_add(headers)
            .ok_or_else(|| write::invalid_input("property record size overflow"))?;
        self.properties
            .iter()
            .try_fold(base, |size, (_, value)| match value {
                BuiltValue::Complex(data) => {
                    i32::try_from(data.len()).map_err(|_err| {
                        write::invalid_input("complex property exceeds i32::MAX")
                    })?;
                    size.checked_add(data.len())
                        .ok_or_else(|| write::invalid_input("property data size overflow"))
                },
                BuiltValue::Simple(_) | BuiltValue::Blip(_) => Ok(size),
            })
    }

    /// Encodes the complete Opt record.
    ///
    /// # Errors
    ///
    /// Returns an error if the property count exceeds twelve bits, if a complex
    /// property exceeds `i32::MAX` bytes, if the record size overflows, or if
    /// writing to `writer` fails.
    pub fn write<W: Write>(&self, writer: &mut W, limits: Limits) -> io::Result<()> {
        let count = u16::try_from(self.properties.len())
            .map_err(|_err| write::invalid_input("too many OfficeArt properties"))?;
        if count > 0x0FFF {
            return Err(write::invalid_input("property count exceeds twelve bits"));
        }
        let total = self.size()?;
        let body_len = total
            .checked_sub(8)
            .ok_or_else(|| write::invalid_input("property record size underflow"))?;
        let body_len = write::bounded_wire_len(body_len, limits, "property payload")?;
        write::atom_header(writer, count, Atom::Opt, body_len, limits)?;

        for (id, value) in &self.properties {
            let (raw_id, raw_value) = match value {
                BuiltValue::Simple(simple) => (*id, *simple),
                BuiltValue::Blip(blip) => (*id | write::BLIP_ID, *blip),
                BuiltValue::Complex(data) => {
                    let len = i32::try_from(data.len()).map_err(|_err| {
                        write::invalid_input("complex property exceeds i32::MAX")
                    })?;
                    (*id | write::COMPLEX, len)
                },
            };
            writer.write_all(&raw_id.to_le_bytes())?;
            writer.write_all(&raw_value.to_le_bytes())?;
        }
        for (_, value) in &self.properties {
            if let BuiltValue::Complex(data) = value {
                writer.write_all(data)?;
            }
        }
        Ok(())
    }
}

/// Builder for one `OfficeArt` shape atom.
#[derive(Debug, Clone, Copy)]
pub struct ShapeBuilder {
    kind: Native,
    id: u32,
    flags: Flags,
}

impl ShapeBuilder {
    /// Creates a shape builder with no flags.
    pub fn new<K: Into<Native>>(kind: K, id: u32) -> Self {
        Self {
            kind: kind.into(),
            id,
            flags: Flags::empty(),
        }
    }

    /// Sets typed shape flags.
    #[must_use]
    pub fn with_flags<F: Into<Flags>>(mut self, flags: F) -> Self {
        self.flags = flags.into();
        self
    }

    /// Encodes the shape atom.
    ///
    /// # Errors
    ///
    /// Returns an error if writing to `writer` fails.
    pub fn write<W: Write>(&self, writer: &mut W, limits: Limits) -> io::Result<()> {
        write::bounded_len(8, limits, "shape payload")?;
        write::atom_header(writer, self.kind.raw(), Atom::Sp, 8, limits)?;
        writer.write_all(Sp::new(self.id, self.flags).as_bytes())
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;
    use crate::{Record, RecordKind};

    #[test]
    fn property_builder_moves_or_borrows_complex_data_and_round_trips() {
        let borrowed = [9, 8, 7];
        let moved = vec![1, 2, 3, 4];
        let moved_pointer = moved.as_ptr();
        let mut builder = PropertyBuilder::new();
        builder.add_blip_id(Id::BlipToDisplay, 7);
        builder.add_complex(Id::Vertices, moved);
        builder.add_complex(Id::SegmentInfo, &borrowed[..]);

        match &builder.properties[1].1 {
            BuiltValue::Complex(Cow::Owned(data)) => assert_eq!(data.as_ptr(), moved_pointer),
            value @ (BuiltValue::Simple(_) | BuiltValue::Blip(_) | BuiltValue::Complex(_)) => {
                panic!("expected moved complex bytes, got {value:?}");
            },
        }
        match &builder.properties[2].1 {
            BuiltValue::Complex(Cow::Borrowed(data)) => {
                assert!(core::ptr::eq(data.as_ptr(), borrowed.as_ptr()));
            },
            value @ (BuiltValue::Simple(_) | BuiltValue::Blip(_) | BuiltValue::Complex(_)) => {
                panic!("expected borrowed complex bytes, got {value:?}");
            },
        }

        let mut bytes = Vec::new();
        builder
            .write(&mut bytes, Limits::default())
            .expect("write properties");
        assert_eq!(builder.size().expect("representable size"), 33);
        assert_eq!(u16::from_le_bytes([bytes[8], bytes[9]]), 0x4104);
        assert_eq!(u16::from_le_bytes([bytes[14], bytes[15]]), 0x8145);
        assert_eq!(u16::from_le_bytes([bytes[20], bytes[21]]), 0x8146);
        assert_eq!(&bytes[26..], &[1, 2, 3, 4, 9, 8, 7]);

        let (record, consumed) =
            Record::parse(&bytes, 0, Limits::default()).expect("parse emitted record");
        assert_eq!(record.kind(), RecordKind::Opt);
        assert_eq!(consumed, bytes.len());
    }

    #[test]
    fn shape_builder_accepts_only_typed_flags() {
        let mut bytes = Vec::new();
        ShapeBuilder::new(Native::RECTANGLE, 42)
            .with_flags(Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
            .write(&mut bytes, Limits::default())
            .expect("write shape");

        let (record, _) = Record::parse(&bytes, 0, Limits::default()).expect("parse shape");
        assert_eq!(record.kind(), RecordKind::Sp);
        assert_eq!(record.instance(), Native::RECTANGLE.raw());
    }
}
