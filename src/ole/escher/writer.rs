//! Escher record writing utilities.
//!
//! Provides helper functions for writing Escher records to binary format.
//! Based on MS-ODRAW specification.

use bitflags::bitflags;
use std::io::{self, Write};
use zerocopy::IntoBytes;
use zerocopy_derive::*;

// =============================================================================
// Shape Flags (MS-ODRAW 2.2.40)
// =============================================================================

bitflags! {
    /// Shape flags for EscherSpRecord (MS-ODRAW 2.2.40)
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub struct ShapeFlags: u32 {
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

pub mod record_type {
    pub const DGG_CONTAINER: u16 = 0xF000;
    pub const BSTORE_CONTAINER: u16 = 0xF001;
    pub const DG_CONTAINER: u16 = 0xF002;
    pub const SPGR_CONTAINER: u16 = 0xF003;
    pub const SP_CONTAINER: u16 = 0xF004;
    pub const DGG: u16 = 0xF006;
    pub const DG: u16 = 0xF008;
    pub const SPGR: u16 = 0xF009;
    pub const SP: u16 = 0xF00A;
    pub const OPT: u16 = 0xF00B;
    pub const CLIENT_ANCHOR: u16 = 0xF010;
    pub const CLIENT_DATA: u16 = 0xF011;
    pub const CLIENT_TEXTBOX: u16 = 0xF00D;
    pub const SPLIT_MENU_COLORS: u16 = 0xF11E;
}

// =============================================================================
// Shape Type Constants (MS-ODRAW 2.4.6 MSOSPT)
// =============================================================================

pub mod shape_type {
    pub const NOT_PRIMITIVE: u16 = 0;
    pub const RECTANGLE: u16 = 1;
    pub const ROUND_RECTANGLE: u16 = 2;
    pub const ELLIPSE: u16 = 3;
    pub const DIAMOND: u16 = 4;
    pub const LINE: u16 = 20;
    pub const TEXT_BOX: u16 = 202;
}

// =============================================================================
// Property Value Constants
// =============================================================================

pub mod prop_value {
    pub const SCHEME_COLOR: u32 = 0x0800_0000;
    pub const SCHEME_FILL: u32 = SCHEME_COLOR | 0x04;
    pub const SCHEME_FILL_BACK: u32 = SCHEME_COLOR;
    pub const SCHEME_LINE: u32 = SCHEME_COLOR | 0x01;
    pub const SCHEME_SHADOW: u32 = SCHEME_COLOR | 0x02;
    pub const LINE_STYLE_DEFAULT: u32 = 0x0010_0010;
    pub const SHAPE_BOOL_DEFAULT: u32 = 0x0008_0008;
}

// =============================================================================
// Zerocopy Data Structures
// =============================================================================

/// Escher record header (8 bytes) - zerocopy compatible
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct EscherRecordHeader {
    pub ver_inst: u16,
    pub rec_type: u16,
    pub length: u32,
}

impl EscherRecordHeader {
    pub const fn new(version: u8, instance: u16, rec_type: u16, length: u32) -> Self {
        let ver_inst = (version as u16 & 0x0F) | ((instance & 0x0FFF) << 4);
        Self {
            ver_inst,
            rec_type,
            length,
        }
    }

    pub const fn container(rec_type: u16, length: u32) -> Self {
        Self::new(0x0F, 0, rec_type, length)
    }
}

/// Shape record data (8 bytes)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C)]
pub struct EscherSpData {
    pub spid: u32,
    pub flags: u32,
}

impl EscherSpData {
    pub const fn new(spid: u32, flags: u32) -> Self {
        Self { spid, flags }
    }

    pub const fn with_flags(spid: u32, flags: ShapeFlags) -> Self {
        Self {
            spid,
            flags: flags.bits(),
        }
    }
}

/// Property entry (6 bytes)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct EscherProperty {
    pub prop_id: u16,
    pub value: u32,
}

impl EscherProperty {
    pub const fn new(prop_id: u16, value: u32) -> Self {
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
pub fn write_record_header<W: Write>(
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
pub fn write_container<W: Write>(
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
pub fn write_atom<W: Write>(
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
pub struct PropertyBuilder {
    properties: Vec<(u16, i32)>,
    complex_data: Vec<u8>,
}

impl PropertyBuilder {
    pub fn new() -> Self {
        Self {
            properties: Vec::new(),
            complex_data: Vec::new(),
        }
    }

    /// Add a simple property.
    pub fn add_simple(&mut self, property_id: u16, value: i32) {
        self.properties.push((property_id, value));
    }

    /// Add a complex property.
    pub fn add_complex(&mut self, property_id: u16, data: &[u8]) {
        let property_id_with_flag = property_id | 0x8000;
        self.properties
            .push((property_id_with_flag, data.len() as i32));
        self.complex_data.extend_from_slice(data);
    }

    /// Write the Opt record.
    pub fn write<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        let num_properties = self.properties.len() as u16;
        let header_size = num_properties as usize * 6;
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
    pub fn size(&self) -> usize {
        8 + (self.properties.len() * 6) + self.complex_data.len()
    }
}

impl Default for PropertyBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Helper to build shape records.
pub struct ShapeBuilder {
    shape_type: u16,
    shape_id: u32,
    flags: u32,
}

impl ShapeBuilder {
    pub fn new(shape_type: u16, shape_id: u32) -> Self {
        Self {
            shape_type,
            shape_id,
            flags: 0,
        }
    }

    pub fn with_flags(mut self, flags: u32) -> Self {
        self.flags = flags;
        self
    }

    /// Write the Sp record.
    pub fn write<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        write_record_header(writer, 0x02, self.shape_type, 0xF00A, 8)?;
        writer.write_all(&self.shape_id.to_le_bytes())?;
        writer.write_all(&self.flags.to_le_bytes())?;
        Ok(())
    }
}

/// Write a ClientAnchor record.
pub fn write_client_anchor<W: Write>(
    writer: &mut W,
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
) -> io::Result<()> {
    write_record_header(writer, 0x00, 0, 0xF010, 16)?;
    writer.write_all(&left.to_le_bytes())?;
    writer.write_all(&top.to_le_bytes())?;
    writer.write_all(&right.to_le_bytes())?;
    writer.write_all(&bottom.to_le_bytes())?;
    Ok(())
}

/// Write a ChildAnchor record.
pub fn write_child_anchor<W: Write>(
    writer: &mut W,
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
) -> io::Result<()> {
    write_record_header(writer, 0x00, 0, 0xF00F, 16)?;
    writer.write_all(&left.to_le_bytes())?;
    writer.write_all(&top.to_le_bytes())?;
    writer.write_all(&right.to_le_bytes())?;
    writer.write_all(&bottom.to_le_bytes())?;
    Ok(())
}

/// Write an Spgr record (group shape coordinates).
pub fn write_spgr<W: Write>(
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

/// Write a Dg record (drawing atom).
pub fn write_dg<W: Write>(writer: &mut W, num_shapes: u32, last_shape_id: u32) -> io::Result<()> {
    write_record_header(writer, 0x00, 0, 0xF008, 8)?;
    writer.write_all(&num_shapes.to_le_bytes())?;
    writer.write_all(&last_shape_id.to_le_bytes())?;
    Ok(())
}
