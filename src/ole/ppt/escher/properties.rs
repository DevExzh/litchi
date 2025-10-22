//! Escher shape property parsing (Opt record).
//!
//! Properties control shape appearance: position, size, colors, rotation, etc.
//! Based on MS-ODRAW specification section 2.3.
//!
//! # Complex Properties
//!
//! Properties can be simple (4-byte value) or complex (variable-length data).
//! Complex properties use a two-pass parsing approach:
//! 1. First pass: Parse all 6-byte property headers
//! 2. Second pass: Read complex data that follows the headers
//!
//! # Performance
//!
//! - Two-pass parsing minimizes data copying
//! - Zero-copy for complex data (borrows from source)
//! - HashMap for O(1) property lookup
//! - Pre-allocated capacity based on property count

use super::container::EscherContainer;
use super::record::EscherRecord;
use super::types::EscherRecordType;
use std::collections::HashMap;

// Property ID flags (from MS-ODRAW)
const IS_BLIP: u16 = 0x4000; // Bit 14: is blip ID
const IS_COMPLEX: u16 = 0x8000; // Bit 15: is complex property
const PROPERTY_ID_MASK: u16 = 0x3FFF; // Lower 14 bits: property number

/// Escher shape property IDs (from MS-ODRAW).
///
/// This enum represents common property IDs used in Office drawings.
/// The actual property number is stored in the lower 14 bits of the property ID.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum EscherPropertyId {
    // Transform properties (0x0000 - 0x003F)
    /// Rotation angle (16.16 fixed point degrees)
    Rotation = 0x0004,

    // Protection properties (0x0040 - 0x007F)
    /// Lock rotation
    LockRotation = 0x0077,
    /// Lock aspect ratio
    LockAspectRatio = 0x0078,
    /// Lock position
    LockPosition = 0x007A,
    /// Lock against grouping
    LockAgainstGrouping = 0x007F,

    // Text properties (0x0080 - 0x00BF)
    /// Text ID (reference to text)
    TextId = 0x0080,
    /// Text left (host margins)
    TextLeft = 0x0081,
    /// Text top (host margins)
    TextTop = 0x0082,
    /// Text right (host margins)
    TextRight = 0x0083,
    /// Text bottom (host margins)
    TextBottom = 0x0084,
    /// Wrap text
    WrapText = 0x0085,
    /// Bidirectional text
    TextBidi = 0x008B,

    // Geometry properties (0x0140 - 0x017F)
    /// Geometry left
    GeomLeft = 0x0140,
    /// Geometry top
    GeomTop = 0x0141,
    /// Geometry right
    GeomRight = 0x0142,
    /// Geometry bottom
    GeomBottom = 0x0143,
    /// Shape path (complex)
    ShapePath = 0x0144,
    /// Vertices (complex array)
    Vertices = 0x0145,
    /// Segment info (complex array)
    SegmentInfo = 0x0146,
    /// Adjust values (complex array)
    AdjustValue = 0x0147,

    // Fill properties (0x0180 - 0x01BF)
    /// Fill type
    FillType = 0x0180,
    /// Fill color (RGB)
    FillColor = 0x0181,
    /// Fill opacity
    FillOpacity = 0x0182,
    /// Fill back color (RGB)
    FillBackColor = 0x0183,
    /// Fill blip (picture)
    FillBlip = 0x0186,
    /// Fill blip name
    FillBlipName = 0x0187,
    /// Fill angle
    FillAngle = 0x018B,
    /// Fill focus
    FillFocus = 0x018C,

    // Line properties (0x01C0 - 0x01FF)
    /// Line color
    LineColor = 0x01C0,
    /// Line style
    LineStyle = 0x01CD,
    /// Line width
    LineWidth = 0x01CB,
    /// Line dash style
    LineDashing = 0x01CE,
    /// Line start arrow head
    LineStartArrowhead = 0x01D0,
    /// Line end arrow head
    LineEndArrowhead = 0x01D1,

    // Shadow properties (0x0200 - 0x023F)
    /// Shadow type
    ShadowType = 0x0200,
    /// Shadow color
    ShadowColor = 0x0201,
    /// Shadow opacity
    ShadowOpacity = 0x0204,
    /// Shadow offset X
    ShadowOffsetX = 0x0205,
    /// Shadow offset Y
    ShadowOffsetY = 0x0206,

    // Picture/Blip properties (0x0100 - 0x010F)
    /// Blip (picture) to display (reference)
    BlipToDisplay = 0x0104,
    /// Picture file name
    PictureFileName = 0x0105,
    /// Picture contrast
    PictureContrast = 0x0108,
    /// Picture brightness
    PictureBrightness = 0x0109,
    /// Gamma
    PictureGamma = 0x010A,
    /// Picture ID
    PictureId = 0x010B,

    // Group properties (0x0380 - 0x03BF)
    /// Group name
    GroupName = 0x0380,
    /// Group description
    GroupDescription = 0x0381,

    /// Unknown property
    Unknown = 0xFFFF,
}

impl From<u16> for EscherPropertyId {
    fn from(value: u16) -> Self {
        // Mask off the flags to get the property number
        let prop_num = value & PROPERTY_ID_MASK;

        match prop_num {
            0x0004 => Self::Rotation,
            0x0081 => Self::TextLeft,
            0x0082 => Self::TextTop,
            0x0083 => Self::TextRight,
            0x0084 => Self::TextBottom,
            0x0077 => Self::LockRotation,
            0x0078 => Self::LockAspectRatio,
            0x007A => Self::LockPosition,
            0x007F => Self::LockAgainstGrouping,
            0x0080 => Self::TextId,
            0x0085 => Self::WrapText,
            0x008B => Self::TextBidi,
            0x0140 => Self::GeomLeft,
            0x0141 => Self::GeomTop,
            0x0142 => Self::GeomRight,
            0x0143 => Self::GeomBottom,
            0x0144 => Self::ShapePath,
            0x0145 => Self::Vertices,
            0x0146 => Self::SegmentInfo,
            0x0147 => Self::AdjustValue,
            0x0180 => Self::FillType,
            0x0181 => Self::FillColor,
            0x0182 => Self::FillOpacity,
            0x0183 => Self::FillBackColor,
            0x0186 => Self::FillBlip,
            0x0187 => Self::FillBlipName,
            0x018B => Self::FillAngle,
            0x018C => Self::FillFocus,
            0x01C0 => Self::LineColor,
            0x01CD => Self::LineStyle,
            0x01CB => Self::LineWidth,
            0x01CE => Self::LineDashing,
            0x01D0 => Self::LineStartArrowhead,
            0x01D1 => Self::LineEndArrowhead,
            0x0200 => Self::ShadowType,
            0x0201 => Self::ShadowColor,
            0x0204 => Self::ShadowOpacity,
            0x0205 => Self::ShadowOffsetX,
            0x0206 => Self::ShadowOffsetY,
            0x0380 => Self::GroupName,
            0x0381 => Self::GroupDescription,
            0x0104 => Self::BlipToDisplay,
            0x0105 => Self::PictureFileName,
            0x0108 => Self::PictureContrast,
            0x0109 => Self::PictureBrightness,
            0x010A => Self::PictureGamma,
            0x010B => Self::PictureId,
            _ => Self::Unknown,
        }
    }
}

/// Escher shape property value.
///
/// Properties can be simple (stored in 4 bytes) or complex (variable-length data).
/// Array properties are a special type of complex property with structured data.
#[derive(Debug, Clone)]
pub enum EscherPropertyValue<'data> {
    /// Simple property: 32-bit integer value
    Simple(i32),

    /// Complex property: binary data (zero-copy borrow)
    ///
    /// This is raw binary data stored in the complex part of the property.
    /// The data is borrowed from the original source for efficiency.
    Complex(&'data [u8]),

    /// Array property: structured array data
    ///
    /// Array properties have a 6-byte header followed by element data:
    /// - 2 bytes: number of elements in array
    /// - 2 bytes: number of elements in memory (reserved)
    /// - 2 bytes: size of each element (can be negative, see get_element_size)
    Array(EscherArrayProperty<'data>),
}

/// Escher array property structure.
///
/// Array properties are complex properties with a specific structure:
/// - 6-byte header (element count, reserved, element size)
/// - Variable-length element data
///
/// # Performance
///
/// - Zero-copy: Borrows data from source
/// - Lazy element access via iterator
/// - No intermediate allocations
#[derive(Debug, Clone)]
pub struct EscherArrayProperty<'data> {
    /// Raw array data including header
    data: &'data [u8],
}

impl<'data> EscherArrayProperty<'data> {
    /// Create array property from raw data.
    ///
    /// # Data Format
    ///
    /// - Bytes 0-1: Number of elements in array (unsigned 16-bit)
    /// - Bytes 2-3: Number of elements in memory / reserved (unsigned 16-bit)
    /// - Bytes 4-5: Size of each element (signed 16-bit, see get_element_size)
    /// - Bytes 6+: Element data
    #[inline]
    pub fn new(data: &'data [u8]) -> Option<Self> {
        if data.len() < 6 {
            return None;
        }
        Some(Self { data })
    }

    /// Get number of elements in array.
    #[inline]
    pub fn element_count(&self) -> u16 {
        if self.data.len() < 2 {
            return 0;
        }
        u16::from_le_bytes([self.data[0], self.data[1]])
    }

    /// Get number of elements in memory (reserved field).
    #[inline]
    pub fn element_count_in_memory(&self) -> u16 {
        if self.data.len() < 4 {
            return 0;
        }
        u16::from_le_bytes([self.data[2], self.data[3]])
    }

    /// Get raw element size value (can be negative).
    #[inline]
    pub fn raw_element_size(&self) -> i16 {
        if self.data.len() < 6 {
            return 0;
        }
        i16::from_le_bytes([self.data[4], self.data[5]])
    }

    /// Get actual element size in bytes.
    ///
    /// # Special Handling
    ///
    /// From MS-ODRAW: If the size is negative, the actual size is:
    /// `(-size) >> 2` (negate and right shift by 2)
    ///
    /// This weird encoding is used for some array properties.
    #[inline]
    pub fn element_size(&self) -> usize {
        let size = self.raw_element_size();
        if size < 0 {
            ((-size) >> 2) as usize
        } else {
            size as usize
        }
    }

    /// Get element at index (zero-copy).
    ///
    /// Returns None if index is out of bounds or element data is truncated.
    #[inline]
    pub fn get_element(&self, index: usize) -> Option<&'data [u8]> {
        let count = self.element_count() as usize;
        if index >= count {
            return None;
        }

        let elem_size = self.element_size();
        let start = 6 + index * elem_size;
        let end = start + elem_size;

        if end > self.data.len() {
            return None;
        }

        Some(&self.data[start..end])
    }

    /// Iterate over all elements (zero-copy).
    pub fn elements(&self) -> impl Iterator<Item = &'data [u8]> {
        let count = self.element_count() as usize;
        let elem_size = self.element_size();
        let data = self.data;

        (0..count).filter_map(move |i| {
            let start = 6 + i * elem_size;
            let end = start + elem_size;
            if end <= data.len() {
                Some(&data[start..end])
            } else {
                None
            }
        })
    }

    /// Get raw array data (including header).
    #[inline]
    pub fn raw_data(&self) -> &'data [u8] {
        self.data
    }
}

/// Escher shape properties collection.
///
/// # Performance
///
/// - HashMap for O(1) property lookup
/// - Pre-allocated capacity based on property count
/// - Two-pass parsing for efficient complex data handling
/// - Zero-copy for complex and array properties
#[derive(Debug, Clone)]
pub struct EscherProperties<'data> {
    properties: HashMap<EscherPropertyId, EscherPropertyValue<'data>>,
}

/// Intermediate property descriptor used during two-pass parsing.
#[derive(Debug, Clone, Copy)]
struct PropertyDescriptor {
    id: EscherPropertyId,
    id_raw: u16,
    value: i32,
}

impl<'data> EscherProperties<'data> {
    /// Create empty properties collection.
    #[inline]
    pub fn new() -> Self {
        Self {
            properties: HashMap::new(),
        }
    }

    /// Parse properties from Escher Opt record using two-pass approach.
    ///
    /// # Algorithm (based on Apache POI)
    ///
    /// **Pass 1**: Parse all 6-byte property headers
    /// - Read property ID (2 bytes) with flags
    /// - Read value/length (4 bytes)
    /// - Create property descriptors
    ///
    /// **Pass 2**: Read complex data for complex properties
    /// - Complex data follows all headers sequentially
    /// - Use length from Pass 1 to read correct amount
    /// - Distinguish between complex and array properties
    ///
    /// # Performance
    ///
    /// - Pre-allocated HashMap capacity
    /// - Zero-copy for complex data (borrows from opt.data)
    /// - Efficient bit manipulation for flags
    /// - Single allocation for property descriptors
    pub fn from_opt_record(opt: &EscherRecord<'data>) -> Self {
        let num_properties = opt.instance as usize;
        let mut properties = HashMap::with_capacity(num_properties);

        if opt.data.len() < 6 {
            return Self { properties };
        }

        // Pass 1: Parse all property headers (6 bytes each)
        let header_size = num_properties * 6;
        if header_size > opt.data.len() {
            // Truncated data, parse what we can
            return Self { properties };
        }

        let mut descriptors = Vec::with_capacity(num_properties);
        for i in 0..num_properties {
            let offset = i * 6;
            if offset + 6 > opt.data.len() {
                break;
            }

            // Read property ID with flags (2 bytes)
            let id_raw = u16::from_le_bytes([opt.data[offset], opt.data[offset + 1]]);

            // Read value/length (4 bytes)
            let value = i32::from_le_bytes([
                opt.data[offset + 2],
                opt.data[offset + 3],
                opt.data[offset + 4],
                opt.data[offset + 5],
            ]);

            let id = EscherPropertyId::from(id_raw);

            descriptors.push(PropertyDescriptor { id, id_raw, value });
        }

        // Pass 2: Process properties and read complex data
        let mut complex_data_offset = header_size;

        for desc in descriptors {
            let is_complex = (desc.id_raw & IS_COMPLEX) != 0;
            let _is_blip = (desc.id_raw & IS_BLIP) != 0;

            let prop_value = if is_complex {
                // Complex property: value is the length of complex data
                let complex_len = desc.value as usize;
                let complex_end = complex_data_offset + complex_len;

                if complex_end > opt.data.len() {
                    // Truncated complex data, skip this property
                    // But still advance offset for next properties
                    complex_data_offset = complex_end;
                    continue;
                }

                let complex_data = &opt.data[complex_data_offset..complex_end];
                complex_data_offset = complex_end;

                // Try to detect array properties by checking if they have array structure
                if Self::is_array_property(complex_data) {
                    if let Some(array_prop) = EscherArrayProperty::new(complex_data) {
                        EscherPropertyValue::Array(array_prop)
                    } else {
                        EscherPropertyValue::Complex(complex_data)
                    }
                } else {
                    EscherPropertyValue::Complex(complex_data)
                }
            } else {
                // Simple property: value is the data itself
                EscherPropertyValue::Simple(desc.value)
            };

            properties.insert(desc.id, prop_value);
        }

        Self { properties }
    }

    /// Heuristic to detect if complex property is an array property.
    ///
    /// Array properties have a 6-byte header followed by array data.
    /// We check if:
    /// 1. Data is at least 6 bytes
    /// 2. Element count and size are reasonable
    /// 3. Total size matches array structure
    fn is_array_property(data: &[u8]) -> bool {
        if data.len() < 6 {
            return false;
        }

        let num_elements = u16::from_le_bytes([data[0], data[1]]) as usize;
        let element_size_raw = i16::from_le_bytes([data[4], data[5]]);

        // Compute actual element size
        let element_size = if element_size_raw < 0 {
            ((-element_size_raw) >> 2) as usize
        } else {
            element_size_raw as usize
        };

        // Check if the data size matches array structure
        // Some arrays don't include header in size calculation, so check both
        let expected_size_with_header = 6 + num_elements * element_size;
        let expected_size_without_header = num_elements * element_size;

        data.len() == expected_size_with_header || data.len() == expected_size_without_header
    }

    /// Parse properties from a container by finding Opt record.
    pub fn from_container(container: &EscherContainer<'data>) -> Self {
        if let Some(opt) = container.find_child(EscherRecordType::Opt) {
            Self::from_opt_record(&opt)
        } else {
            Self::new()
        }
    }

    /// Get property value by ID.
    #[inline]
    pub fn get(&self, id: EscherPropertyId) -> Option<&EscherPropertyValue<'data>> {
        self.properties.get(&id)
    }

    /// Get simple integer property value.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(rotation) = props.get_int(EscherPropertyId::Rotation) {
    ///     println!("Rotation: {}", rotation);
    /// }
    /// ```
    #[inline]
    pub fn get_int(&self, id: EscherPropertyId) -> Option<i32> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Simple(v)) => Some(*v),
            _ => None,
        }
    }

    /// Get color property value (RGB).
    ///
    /// Colors are stored as 32-bit values in BGR format (low byte = blue).
    #[inline]
    pub fn get_color(&self, id: EscherPropertyId) -> Option<u32> {
        self.get_int(id).map(|v| v as u32)
    }

    /// Get boolean property value.
    ///
    /// Interprets non-zero values as true, zero as false.
    #[inline]
    pub fn get_bool(&self, id: EscherPropertyId) -> Option<bool> {
        self.get_int(id).map(|v| v != 0)
    }

    /// Get complex binary property value (zero-copy).
    ///
    /// Returns borrowed slice from original data.
    #[inline]
    pub fn get_binary(&self, id: EscherPropertyId) -> Option<&'data [u8]> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Complex(data)) => Some(data),
            _ => None,
        }
    }

    /// Get array property value.
    ///
    /// # Example
    ///
    /// ```ignore
    /// if let Some(vertices) = props.get_array(EscherPropertyId::Vertices) {
    ///     for vertex in vertices.elements() {
    ///         // Process vertex data
    ///     }
    /// }
    /// ```
    #[inline]
    pub fn get_array(&self, id: EscherPropertyId) -> Option<&EscherArrayProperty<'data>> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Array(array)) => Some(array),
            _ => None,
        }
    }

    /// Check if property exists.
    #[inline]
    pub fn has(&self, id: EscherPropertyId) -> bool {
        self.properties.contains_key(&id)
    }

    /// Get number of properties.
    #[inline]
    pub fn len(&self) -> usize {
        self.properties.len()
    }

    /// Check if properties collection is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    /// Iterate over all properties.
    pub fn iter(&self) -> impl Iterator<Item = (&EscherPropertyId, &EscherPropertyValue<'data>)> {
        self.properties.iter()
    }
}

impl<'data> Default for EscherProperties<'data> {
    fn default() -> Self {
        Self::new()
    }
}

/// Shape anchor (position and size).
///
/// # Coordinates
///
/// - Coordinates are in master units (typically 1/576 inch)
/// - Origin is top-left corner
#[derive(Debug, Clone, Copy)]
pub struct ShapeAnchor {
    /// Left coordinate
    pub left: i32,
    /// Top coordinate
    pub top: i32,
    /// Right coordinate
    pub right: i32,
    /// Bottom coordinate
    pub bottom: i32,
}

impl ShapeAnchor {
    /// Create anchor from coordinates.
    #[inline]
    pub const fn new(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    /// Get width.
    #[inline]
    pub const fn width(&self) -> i32 {
        self.right - self.left
    }

    /// Get height.
    #[inline]
    pub const fn height(&self) -> i32 {
        self.bottom - self.top
    }

    /// Parse from ChildAnchor record.
    pub fn from_child_anchor(anchor: &EscherRecord) -> Option<Self> {
        if anchor.data.len() < 16 {
            return None;
        }

        let left = i32::from_le_bytes([
            anchor.data[0],
            anchor.data[1],
            anchor.data[2],
            anchor.data[3],
        ]);
        let top = i32::from_le_bytes([
            anchor.data[4],
            anchor.data[5],
            anchor.data[6],
            anchor.data[7],
        ]);
        let right = i32::from_le_bytes([
            anchor.data[8],
            anchor.data[9],
            anchor.data[10],
            anchor.data[11],
        ]);
        let bottom = i32::from_le_bytes([
            anchor.data[12],
            anchor.data[13],
            anchor.data[14],
            anchor.data[15],
        ]);

        Some(Self::new(left, top, right, bottom))
    }

    /// Parse from ClientAnchor record.
    pub fn from_client_anchor(anchor: &EscherRecord) -> Option<Self> {
        // ClientAnchor has same format as ChildAnchor
        Self::from_child_anchor(anchor)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_anchor_dimensions() {
        let anchor = ShapeAnchor::new(100, 200, 500, 600);
        assert_eq!(anchor.width(), 400);
        assert_eq!(anchor.height(), 400);
    }
}
