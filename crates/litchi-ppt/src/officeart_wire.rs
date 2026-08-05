//! Escher record writing utilities.
//!
//! Provides helper functions for writing Escher records to binary format.
//! Based on MS-ODRAW specification.

use bitflags::bitflags;
#[cfg(test)]
use std::io::{self, Write};
#[cfg(test)]
use zerocopy::IntoBytes;
use zerocopy_derive::{Immutable, IntoBytes, KnownLayout};

/// OfficeArtFOPTEOPID `fBid`: the property value is a BLIP identifier.
#[cfg(test)]
const PROPERTY_FLAG_BLIP_ID: u16 = 0x4000;
/// OfficeArtFOPTEOPID `fComplex`: the property value is stored after the table.
pub(crate) const PROPERTY_FLAG_COMPLEX: u16 = 0x8000;

// =============================================================================
// Shape Flags (MS-ODRAW 2.2.40)
// =============================================================================

bitflags! {
    /// Shape flags for EscherSpRecord (MS-ODRAW 2.2.40)
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub(crate) struct ShapeFlags: u32 {
        /// Shape is a group
        const GROUP = 0x0001;
        /// Shape is a child of a group
        const CHILD = 0x0002;
        /// Shape is the topmost group (patriarch)
        const PATRIARCH = 0x0004;
        /// Shape has been deleted
        const DELETED = 0x0008;
        /// Shape is an OLE object
        const OLE_SHAPE = 0x0010;
        /// Shape has a valid master
        const HAVE_MASTER = 0x0020;
        /// Shape is flipped horizontally
        const FLIP_H = 0x0040;
        /// Shape is flipped vertically
        const FLIP_V = 0x0080;
        /// Shape is a connector
        const CONNECTOR = 0x0100;
        /// Shape has an anchor
        const HAVE_ANCHOR = 0x0200;
        /// Shape is a background shape
        const BACKGROUND = 0x0400;
        /// Shape has a shape type property
        const HAVE_SPT = 0x0800;
    }
}

// =============================================================================
// Escher Record Types Constants
// =============================================================================

pub(crate) mod record_type {
    pub(crate) const DGG_CONTAINER: u16 = 0xF000;
    pub(crate) const DG_CONTAINER: u16 = 0xF002;
    pub(crate) const SPGR_CONTAINER: u16 = 0xF003;
    pub(crate) const SP_CONTAINER: u16 = 0xF004;
    pub(crate) const DGG: u16 = 0xF006;
    pub(crate) const DG: u16 = 0xF008;
    pub(crate) const SPGR: u16 = 0xF009;
    pub(crate) const SP: u16 = 0xF00A;
    pub(crate) const OPT: u16 = 0xF00B;
    pub(crate) const CLIENT_ANCHOR: u16 = 0xF010;
    pub(crate) const CLIENT_DATA: u16 = 0xF011;
    #[cfg(test)]
    pub(crate) const CLIENT_TEXTBOX: u16 = 0xF00D;
    #[cfg(test)]
    pub(crate) const CHILD_ANCHOR: u16 = 0xF00F;
    pub(crate) const SPLIT_MENU_COLORS: u16 = 0xF11E;
}

// =============================================================================
// Shape Type Constants (MS-ODRAW 2.4.6 MSOSPT)
// =============================================================================

pub(crate) mod shape_type {
    pub(crate) const NOT_PRIMITIVE: u16 = 0;
    pub(crate) const RECTANGLE: u16 = 1;
    pub(crate) const ROUND_RECTANGLE: u16 = 2;
    pub(crate) const ELLIPSE: u16 = 3;
    pub(crate) const DIAMOND: u16 = 4;
    pub(crate) const LINE: u16 = 20;
    pub(crate) const TEXT_BOX: u16 = 202;
}

// =============================================================================
// Property Value Constants
// =============================================================================

pub(crate) mod prop_value {
    pub(crate) const SCHEME_COLOR: u32 = 0x0800_0000;
    pub(crate) const SCHEME_FILL: u32 = SCHEME_COLOR | 0x04;
    pub(crate) const SCHEME_FILL_BACK: u32 = SCHEME_COLOR;
    pub(crate) const SCHEME_LINE: u32 = SCHEME_COLOR | 0x01;
    pub(crate) const SCHEME_SHADOW: u32 = SCHEME_COLOR | 0x02;
    pub(crate) const LINE_STYLE_DEFAULT: u32 = 0x0010_0010;
    pub(crate) const LINE_STYLE_BOOL_DEFAULT: u32 = 0x0008_0008;
    pub(crate) const FILL_STYLE_DISABLED: u32 = 0x0010_0000;
    pub(crate) const FILL_STYLE_ENABLED: u32 = 0x0015_0011;
    pub(crate) const SHADOW_STYLE_DISABLED: u32 = 0x0002_0000;
    pub(crate) const SHADOW_STYLE_ENABLED: u32 = 0x0002_0002;
}

// =============================================================================
// Zerocopy Data Structures
// =============================================================================

/// Escher record header (8 bytes) - zerocopy compatible
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub(crate) struct EscherRecordHeader {
    pub(crate) ver_inst: u16,
    pub(crate) rec_type: u16,
    pub(crate) length: u32,
}

impl EscherRecordHeader {
    pub(crate) const fn new(version: u8, instance: u16, rec_type: u16, length: u32) -> Self {
        let ver_inst = (u16::from(version) & 0x0F) | ((instance & 0x0FFF) << 4);
        Self {
            ver_inst,
            rec_type,
            length,
        }
    }
}

/// Shape record data (8 bytes)
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub(crate) struct EscherSpData {
    pub(crate) spid: u32,
    pub(crate) flags: u32,
}

impl EscherSpData {
    pub(crate) const fn with_flags(spid: u32, flags: ShapeFlags) -> Self {
        Self {
            spid,
            flags: flags.bits(),
        }
    }
}

/// Property entry (6 bytes)
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub(crate) struct EscherProperty {
    pub(crate) prop_id: u16,
    pub(crate) value: u32,
}

impl EscherProperty {
    pub(crate) const fn new(prop_id: u16, value: u32) -> Self {
        Self { prop_id, value }
    }
}

// =============================================================================
// Writing Functions
// =============================================================================

/// Write an Escher record header (8 bytes).
///
/// # Format
///
/// - Bytes 0-1: Version (4 bits) | Instance (12 bits)
/// - Bytes 2-3: Record Type
/// - Bytes 4-7: Record Length (32-bit)
#[cfg(test)]
pub(crate) fn write_record_header<W: Write>(
    writer: &mut W,
    version: u8,
    instance: u16,
    record_type: u16,
    length: u32,
) -> io::Result<()> {
    let header = EscherRecordHeader::new(version, instance, record_type, length);
    writer.write_all(header.as_bytes())?;
    Ok(())
}

/// Write a container record with pre-calculated child data.
#[cfg(test)]
pub(crate) fn write_container<W: Write>(
    writer: &mut W,
    instance: u16,
    record_type: u16,
    child_data: &[u8],
) -> io::Result<()> {
    write_record_header(writer, 0x0F, instance, record_type, child_data.len() as u32)?;
    writer.write_all(child_data)?;
    Ok(())
}

/// Write a simple atom record.
#[cfg(test)]
pub(crate) fn write_atom<W: Write>(
    writer: &mut W,
    version: u8,
    instance: u16,
    record_type: u16,
    data: &[u8],
) -> io::Result<()> {
    write_record_header(writer, version, instance, record_type, data.len() as u32)?;
    writer.write_all(data)?;
    Ok(())
}

/// Helper to build property records (Opt records).
#[cfg(test)]
pub(crate) struct PropertyBuilder {
    properties: Vec<(u16, i32)>,
    complex_data: Vec<u8>,
}

#[cfg(test)]
impl PropertyBuilder {
    pub(crate) fn new() -> Self {
        Self {
            properties: Vec::new(),
            complex_data: Vec::new(),
        }
    }

    /// Add a simple property.
    pub(crate) fn add_simple(&mut self, property_id: u16, value: i32) {
        self.properties.push((property_id, value));
    }

    /// Add a simple property whose value identifies an entry in the BLIP store.
    pub(crate) fn add_blip_id(&mut self, property_id: u16, blip_id: i32) {
        self.properties
            .push((property_id | PROPERTY_FLAG_BLIP_ID, blip_id));
    }

    /// Add a complex property.
    pub(crate) fn add_complex(&mut self, property_id: u16, data: &[u8]) {
        let property_id_with_flag = property_id | PROPERTY_FLAG_COMPLEX;
        self.properties
            .push((property_id_with_flag, data.len() as i32));
        self.complex_data.extend_from_slice(data);
    }

    /// Write the Opt record.
    pub(crate) fn write<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        let num_properties = self.properties.len() as u16;
        let header_size = usize::from(num_properties) * 6;
        let total_size = header_size + self.complex_data.len();

        write_record_header(writer, 0x03, num_properties, 0xF00B, total_size as u32)?;

        for (prop_id, value) in &self.properties {
            writer.write_all(&prop_id.to_le_bytes())?;
            writer.write_all(&value.to_le_bytes())?;
        }

        writer.write_all(&self.complex_data)?;
        Ok(())
    }

    /// Get the total size that would be written.
    pub(crate) fn size(&self) -> usize {
        8 + (self.properties.len() * 6) + self.complex_data.len()
    }
}

#[cfg(test)]
impl Default for PropertyBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Helper to build shape records.
#[cfg(test)]
pub(crate) struct ShapeBuilder {
    shape_type: u16,
    shape_id: u32,
    flags: u32,
}

#[cfg(test)]
impl ShapeBuilder {
    pub(crate) fn new(shape_type: u16, shape_id: u32) -> Self {
        Self {
            shape_type,
            shape_id,
            flags: 0,
        }
    }

    pub(crate) fn with_flags(mut self, flags: u32) -> Self {
        self.flags = flags;
        self
    }

    /// Write the Sp record.
    pub(crate) fn write<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        write_record_header(writer, 0x02, self.shape_type, 0xF00A, 8)?;
        writer.write_all(&self.shape_id.to_le_bytes())?;
        writer.write_all(&self.flags.to_le_bytes())?;
        Ok(())
    }
}

/// Write a ChildAnchor record.
#[cfg(test)]
pub(crate) fn write_child_anchor<W: Write>(
    writer: &mut W,
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
) -> io::Result<()> {
    write_record_header(writer, 0x00, 0, record_type::CHILD_ANCHOR, 16)?;
    writer.write_all(&left.to_le_bytes())?;
    writer.write_all(&top.to_le_bytes())?;
    writer.write_all(&right.to_le_bytes())?;
    writer.write_all(&bottom.to_le_bytes())?;
    Ok(())
}

/// Write an Spgr record (group shape coordinates).
#[cfg(test)]
pub(crate) fn write_spgr<W: Write>(
    writer: &mut W,
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
) -> io::Result<()> {
    write_record_header(writer, 0x01, 0, 0xF009, 16)?;
    writer.write_all(&left.to_le_bytes())?;
    writer.write_all(&top.to_le_bytes())?;
    writer.write_all(&right.to_le_bytes())?;
    writer.write_all(&bottom.to_le_bytes())?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn property_builder_distinguishes_blip_and_complex_flags() {
        let mut builder = PropertyBuilder::new();
        builder.add_blip_id(0x0104, 7);
        builder.add_complex(0x0145, &[1, 2, 3, 4]);

        let mut bytes = Vec::new();
        builder.write(&mut bytes).unwrap();

        assert_eq!(builder.size(), 24);
        assert_eq!(u16::from_le_bytes([bytes[8], bytes[9]]), 0x4104);
        assert_eq!(i32::from_le_bytes(bytes[10..14].try_into().unwrap()), 7);
        assert_eq!(u16::from_le_bytes([bytes[14], bytes[15]]), 0x8145);
        assert_eq!(i32::from_le_bytes(bytes[16..20].try_into().unwrap()), 4);
        assert_eq!(&bytes[20..], &[1, 2, 3, 4]);
    }
}
