//! Escher shape property parsing (Opt record).
//!
//! Properties control shape appearance: position, size, colors, rotation, etc.
//! Based on MS-ODRAW specification section 2.3.

use super::record::EscherRecord;
use super::types::EscherRecordType;
use super::container::EscherContainer;
use std::collections::HashMap;

/// Escher shape property IDs (from MS-ODRAW).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum EscherPropertyId {
    // Transform properties
    /// Rotation angle (16.16 fixed point degrees)
    Rotation = 0x0004,
    
    // Geometry properties
    /// Left coordinate
    Left = 0x0082,
    /// Top coordinate
    Top = 0x0083,
    /// Right coordinate
    Right = 0x0084,
    /// Bottom coordinate
    Bottom = 0x0085,
    
    // Fill properties
    /// Fill color
    FillColor = 0x0181,
    /// Fill type
    FillType = 0x0182,
    /// Fill opacity
    FillOpacity = 0x0183,
    
    // Line properties
    /// Line color
    LineColor = 0x01C0,
    /// Line width
    LineWidth = 0x01CB,
    /// Line style
    LineStyle = 0x01C9,
    
    // Text properties
    /// Text ID (reference to text)
    TextId = 0x0080,
    /// Text left margin
    TextLeftMargin = 0x0065,
    /// Text top margin
    TextTopMargin = 0x0066,
    /// Text right margin
    TextRightMargin = 0x0067,
    /// Text bottom margin
    TextBottomMargin = 0x0068,
    
    // Picture properties
    /// Blip (picture) reference
    PictureId = 0x0104,
    /// Picture file name
    PictureName = 0x0105,
    
    // Protection properties
    /// Lock rotation
    LockRotation = 0x0077,
    /// Lock aspect ratio
    LockAspectRatio = 0x0078,
    
    /// Unknown property
    Unknown = 0xFFFF,
}

impl From<u16> for EscherPropertyId {
    fn from(value: u16) -> Self {
        match value {
            0x0004 => Self::Rotation,
            0x0082 => Self::Left,
            0x0083 => Self::Top,
            0x0084 => Self::Right,
            0x0085 => Self::Bottom,
            0x0181 => Self::FillColor,
            0x0182 => Self::FillType,
            0x0183 => Self::FillOpacity,
            0x01C0 => Self::LineColor,
            0x01C9 => Self::LineStyle,
            0x01CB => Self::LineWidth,
            0x0080 => Self::TextId,
            0x0065 => Self::TextLeftMargin,
            0x0066 => Self::TextTopMargin,
            0x0067 => Self::TextRightMargin,
            0x0068 => Self::TextBottomMargin,
            0x0104 => Self::PictureId,
            0x0105 => Self::PictureName,
            0x0077 => Self::LockRotation,
            0x0078 => Self::LockAspectRatio,
            _ => Self::Unknown,
        }
    }
}

/// Escher shape property value.
#[derive(Debug, Clone)]
pub enum EscherPropertyValue {
    /// 32-bit integer value
    Integer(i32),
    /// Boolean value
    Boolean(bool),
    /// Color value (RGB)
    Color(u32),
    /// Binary data (for complex properties)
    Binary(Vec<u8>),
}

/// Escher shape properties collection.
///
/// # Performance
///
/// - HashMap for O(1) property lookup
/// - Pre-allocated capacity based on property count
#[derive(Debug, Clone)]
pub struct EscherProperties {
    properties: HashMap<EscherPropertyId, EscherPropertyValue>,
}

impl EscherProperties {
    /// Create empty properties collection.
    #[inline]
    pub fn new() -> Self {
        Self {
            properties: HashMap::new(),
        }
    }
    
    /// Parse properties from Escher Opt record.
    ///
    /// # Performance
    ///
    /// - Single-pass parsing
    /// - Pre-allocated HashMap capacity
    /// - Efficient bit manipulation
    pub fn from_opt_record(opt: &EscherRecord) -> Self {
        let mut properties = HashMap::with_capacity(opt.instance as usize);
        
        if opt.data.len() < 6 {
            return Self { properties };
        }
        
        // Each property is 6 bytes: 2 bytes ID + 4 bytes value
        let mut offset = 0;
        while offset + 6 <= opt.data.len() {
            let prop_id_raw = u16::from_le_bytes([opt.data[offset], opt.data[offset + 1]]);
            let prop_id = EscherPropertyId::from(prop_id_raw & 0x3FFF); // Lower 14 bits
            let is_complex = (prop_id_raw & 0x8000) != 0; // Bit 15
            
            let value_bytes = [
                opt.data[offset + 2],
                opt.data[offset + 3],
                opt.data[offset + 4],
                opt.data[offset + 5],
            ];
            let value = i32::from_le_bytes(value_bytes);
            
            let prop_value = if is_complex {
                // Complex property - value is length, data follows
                EscherPropertyValue::Binary(Vec::new()) // TODO: Parse complex data
            } else {
                // Simple property - value is the data
                EscherPropertyValue::Integer(value)
            };
            
            properties.insert(prop_id, prop_value);
            offset += 6;
        }
        
        Self { properties }
    }
    
    /// Parse properties from a container by finding Opt record.
    pub fn from_container(container: &EscherContainer) -> Self {
        if let Some(opt) = container.find_child(EscherRecordType::Opt) {
            Self::from_opt_record(&opt)
        } else {
            Self::new()
        }
    }
    
    /// Get integer property value.
    #[inline]
    pub fn get_int(&self, id: EscherPropertyId) -> Option<i32> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Integer(v)) => Some(*v),
            _ => None,
        }
    }
    
    /// Get color property value (RGB).
    #[inline]
    pub fn get_color(&self, id: EscherPropertyId) -> Option<u32> {
        self.get_int(id).map(|v| v as u32)
    }
    
    /// Get boolean property value.
    #[inline]
    pub fn get_bool(&self, id: EscherPropertyId) -> Option<bool> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Boolean(v)) => Some(*v),
            Some(EscherPropertyValue::Integer(v)) => Some(*v != 0),
            _ => None,
        }
    }
    
    /// Get binary property value.
    #[inline]
    pub fn get_binary(&self, id: EscherPropertyId) -> Option<&[u8]> {
        match self.properties.get(&id) {
            Some(EscherPropertyValue::Binary(v)) => Some(v),
            _ => None,
        }
    }
    
    /// Check if property exists.
    #[inline]
    pub fn has(&self, id: EscherPropertyId) -> bool {
        self.properties.contains_key(&id)
    }
}

impl Default for EscherProperties {
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
        Self { left, top, right, bottom }
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
            anchor.data[0], anchor.data[1], anchor.data[2], anchor.data[3]
        ]);
        let top = i32::from_le_bytes([
            anchor.data[4], anchor.data[5], anchor.data[6], anchor.data[7]
        ]);
        let right = i32::from_le_bytes([
            anchor.data[8], anchor.data[9], anchor.data[10], anchor.data[11]
        ]);
        let bottom = i32::from_le_bytes([
            anchor.data[12], anchor.data[13], anchor.data[14], anchor.data[15]
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

