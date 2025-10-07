/// Escher record parsing for PowerPoint shapes.
///
/// This module provides functionality to parse Escher binary records
/// that contain shape data in PowerPoint presentations.
///
/// Escher is Microsoft's binary format for storing graphics and shape data
/// in Office documents, including PowerPoint presentations.
use super::shape::{ShapeProperties, ShapeType};
use super::super::package::{PptError, Result};
use std::collections::HashMap;

/// Escher property types for Office Drawing properties
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum EscherPropertyType {
    /// Transform properties
    Transform = 0x0000,
    /// Fill style properties
    FillStyle = 0x0001,
    /// Line style properties
    LineStyle = 0x0002,
    /// Shadow style properties
    ShadowStyle = 0x0003,
    /// Geometry properties
    Geometry = 0x0004,
    /// Text properties
    Text = 0x0005,
    /// 3D properties
    Properties3D = 0x0006,
    /// Group shape properties
    GroupShape = 0x0007,
    /// Unknown property type
    Unknown = 0xFFFF,
}

impl From<u16> for EscherPropertyType {
    fn from(value: u16) -> Self {
        match value {
            0x0000 => EscherPropertyType::Transform,
            0x0001 => EscherPropertyType::FillStyle,
            0x0002 => EscherPropertyType::LineStyle,
            0x0003 => EscherPropertyType::ShadowStyle,
            0x0004 => EscherPropertyType::Geometry,
            0x0005 => EscherPropertyType::Text,
            0x0006 => EscherPropertyType::Properties3D,
            0x0007 => EscherPropertyType::GroupShape,
            _ => EscherPropertyType::Unknown,
        }
    }
}

/// Escher property holder types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EscherPropertyHolder {
    /// Simple property (fixed size)
    Simple,
    /// Boolean property
    Boolean,
    /// RGB color property
    RGB,
    /// Shape path property
    ShapePath,
    /// Array property
    Array,
    /// Complex property (variable size)
    Complex,
}

/// An Escher property containing binary data and metadata
#[derive(Debug, Clone)]
pub struct EscherProperty {
    /// Property ID (includes type, complex flag, blip flag)
    pub id: u16,
    /// Property data value
    pub data: u32,
    /// Complex data (for complex properties)
    pub complex_data: Option<Vec<u8>>,
    /// Array data (for array properties)
    pub array_data: Option<Vec<u8>>,
}

impl EscherProperty {
    /// Create a new Escher property
    pub fn new(id: u16, data: u32) -> Self {
        Self {
            id,
            data,
            complex_data: None,
            array_data: None,
        }
    }

    /// Create a complex Escher property
    pub fn new_complex(id: u16, data: u32, complex_data: Vec<u8>) -> Self {
        Self {
            id,
            data,
            complex_data: Some(complex_data),
            array_data: None,
        }
    }

    /// Create an array Escher property
    pub fn new_array(id: u16, data: u32, array_data: Vec<u8>) -> Self {
        Self {
            id,
            data,
            complex_data: None,
            array_data: Some(array_data),
        }
    }

    /// Get the property number (lower 14 bits)
    pub fn property_number(&self) -> u16 {
        self.id & 0x3FFF
    }

    /// Check if this is a complex property
    pub fn is_complex(&self) -> bool {
        (self.id & 0x8000) != 0
    }

    /// Check if this is a blip ID property
    pub fn is_blip_id(&self) -> bool {
        (self.id & 0x4000) != 0
    }

    /// Get the property type based on the property number
    pub fn property_type(&self) -> EscherPropertyType {
        EscherPropertyType::from(self.property_number())
    }

    /// Get the property holder type based on the property number
    pub fn property_holder(&self) -> EscherPropertyHolder {
        // Based on POI's EscherPropertyTypes.forPropertyID logic
        match self.property_number() {
            // Boolean properties (0x00BF - 0x013F)
            0x00BF..=0x013F => EscherPropertyHolder::Boolean,
            // RGB properties (0x0140 - 0x017F)
            0x0140..=0x017F => EscherPropertyHolder::RGB,
            // Shape path properties (0x0180 - 0x01BF)
            0x0180..=0x01BF => EscherPropertyHolder::ShapePath,
            // Array properties (0x01C0 - 0x01FF)
            0x01C0..=0x01FF => EscherPropertyHolder::Array,
            // Complex properties (0x0200 - 0x03FF)
            0x0200..=0x03FF => EscherPropertyHolder::Complex,
            // Simple properties (everything else)
            _ => EscherPropertyHolder::Simple,
        }
    }

    /// Parse properties from binary data (based on POI's EscherPropertyFactory).
    /// Optimized for performance with pre-allocation and minimal copying.
    pub fn parse_properties(data: &[u8], num_properties: u16) -> Result<Vec<Self>> {
        if num_properties == 0 {
            return Ok(Vec::new());
        }

        // Pre-allocate with exact size to avoid reallocations
        let mut properties = Vec::with_capacity(num_properties as usize);
        let mut offset = 0;

        for _ in 0..num_properties {
            if offset + 6 > data.len() {
                break; // Not enough data for property header
            }

            // Use unsafe for performance - we already checked bounds
            let prop_id = unsafe { u16::from_le_bytes(*(&data[offset..offset + 2] as *const [u8] as *const [u8; 2])) };
            let prop_data = unsafe { u32::from_le_bytes(*(&data[offset + 2..offset + 6] as *const [u8] as *const [u8; 4])) };

            let is_complex = (prop_id & 0x8000) != 0;

            let property = if is_complex {
                // Parse complex property data
                let complex_size = (prop_data >> 16) as usize; // High 16 bits contain size
                if offset + 6 + complex_size > data.len() {
                    break; // Not enough data for complex property
                }
                let complex_data = data[offset + 6..offset + 6 + complex_size].to_vec();
                offset += 6 + complex_size;

                Self::new_complex(prop_id, prop_data, complex_data)
            } else {
                // Simple property - advance offset without additional bounds check
                offset += 6;
                Self::new(prop_id, prop_data)
            };

            properties.push(property);
        }

        // Shrink to fit if we allocated more than needed
        properties.shrink_to_fit();
        Ok(properties)
    }
}

/// Property values extracted from Escher records for convenient access
#[derive(Debug, Clone, Default)]
pub struct PropertyValues {
    // Fill properties
    pub fill_type: Option<u16>,
    pub fill_color: Option<u32>,
    pub fill_opacity: Option<u16>,
    pub fill_back_color: Option<u32>,

    // Line properties
    pub line_color: Option<u32>,
    pub line_opacity: Option<u16>,
    pub line_width: Option<u16>,
    pub line_style: Option<u16>,
    pub line_dash_style: Option<u16>,

    // Shadow properties
    pub shadow_type: Option<u16>,
    pub shadow_color: Option<u32>,
    pub shadow_opacity: Option<u16>,
    pub shadow_offset_x: Option<i32>,
    pub shadow_offset_y: Option<i32>,

    // Text properties
    pub text_left_margin: Option<i32>,
    pub text_top_margin: Option<i32>,
    pub text_right_margin: Option<i32>,
    pub text_bottom_margin: Option<i32>,
    pub text_anchor: Option<u16>,

    // Transform properties
    pub rotation: Option<u32>,
    pub lock_aspect_ratio: Option<bool>,
}

/// Escher record types used in PowerPoint shapes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum EscherRecordType {
    /// Container record (contains other records)
    Container = 0xF000,
    /// Shape properties record
    ShapeProperties = 0xF004,
    /// Text properties record
    TextProperties = 0xF005,
    /// Geometry properties record
    GeometryProperties = 0xF006,
    /// Fill properties record
    FillProperties = 0xF007,
    /// Line properties record
    LineProperties = 0xF008,
    /// Shadow properties record
    ShadowProperties = 0xF009,
    /// Perspective properties record
    PerspectiveProperties = 0xF00A,
    /// 3D properties record
    Properties3D = 0xF00B,
    /// Transform record (position, size, rotation)
    Transform = 0xF010,
    /// Text record (contains text content)
    Text = 0xF011,
    /// Child anchor record
    ChildAnchor = 0xF00C,
    /// Client anchor record
    ClientAnchor = 0xF00D,
    /// Client data record
    ClientData = 0xF00E,
    /// Placeholder data record
    PlaceholderData = 0xF00F,
    /// Options record (properties)
    Options = 0xF122,
}

/// Helper function to get the u16 value of an EscherRecordType
impl EscherRecordType {
    pub fn as_u16(self) -> u16 {
        unsafe { std::mem::transmute::<Self, u16>(self) }
    }
}

impl From<u16> for EscherRecordType {
    fn from(value: u16) -> Self {
        match value {
            0xF000 => EscherRecordType::Container,
            0xF004 => EscherRecordType::ShapeProperties,
            0xF005 => EscherRecordType::TextProperties,
            0xF006 => EscherRecordType::GeometryProperties,
            0xF007 => EscherRecordType::FillProperties,
            0xF008 => EscherRecordType::LineProperties,
            0xF009 => EscherRecordType::ShadowProperties,
            0xF00A => EscherRecordType::PerspectiveProperties,
            0xF00B => EscherRecordType::Properties3D,
            0xF010 => EscherRecordType::Transform,
            0xF011 => EscherRecordType::Text,
            0xF00C => EscherRecordType::ChildAnchor,
            0xF00D => EscherRecordType::ClientAnchor,
            0xF00E => EscherRecordType::ClientData,
            0xF00F => EscherRecordType::PlaceholderData,
            0xF122 => EscherRecordType::Options,
            _ => EscherRecordType::Container, // Default fallback for unknown types
        }
    }
}

/// An Escher record containing binary data and metadata.
/// Optimized for performance with zero-copy parsing where possible.
#[derive(Debug, Clone)]
pub struct EscherRecord {
    /// Record type
    pub record_type: EscherRecordType,
    /// Record version
    pub version: u16,
    /// Record instance (sub-type)
    pub instance: u16,
    /// Record data length
    pub data_length: u32,
    /// Record data (owned for now, could be Cow in future)
    pub data: Vec<u8>,
    /// Child records (for container records)
    pub children: Vec<EscherRecord>,
    /// Parsed properties (for Options records)
    pub properties: Vec<EscherProperty>,
}

/// Actions to take when processing text characters in UTF-16LE decoding.
#[derive(Debug, Clone, Copy)]
enum CharacterAction {
    /// Add the character to the text
    Add(char),
    /// Stop processing (control character encountered)
    Stop,
    /// Skip this character (invalid or unhandled)
    Skip,
}

impl CharacterAction {
    /// Process a UTF-16LE code unit and determine the appropriate action.
    /// This replaces the previous if-else chain with a more idiomatic match expression.
    fn process_text_character(code_unit: u16) -> Self {
            match code_unit {
            // Null terminator - stop processing
            0 => CharacterAction::Stop,
            // Vertical tab (often used as paragraph separator) - stop processing
            0x0B => CharacterAction::Stop,
            // ASCII range (0x01-0x7F) - add as character (excluding 0 and 0x0B)
            0x01..=0x0A | 0x0C..=0x7F => {
                if let Some(ch) = char::from_u32(code_unit as u32) {
                    CharacterAction::Add(ch)
                } else {
                    CharacterAction::Skip
                }
            }
            // Unicode range (0x80 and above) - try to decode as Unicode
            0x80.. => {
                if let Some(ch) = char::from_u32(code_unit as u32) {
                    CharacterAction::Add(ch)
                } else {
                    CharacterAction::Skip
                }
            }
        }
    }
}

impl EscherRecord {
    /// Parse an Escher record from binary data.
    /// Optimized for performance with minimal allocations.
    ///
    /// # Arguments
    ///
    /// * `data` - Binary data containing the record
    /// * `offset` - Starting offset in the data
    ///
    /// # Returns
    ///
    /// Tuple of (parsed_record, bytes_consumed)
    pub fn parse(data: &[u8], offset: usize) -> Result<(Self, usize)> {
        if offset + 8 > data.len() {
            return Err(PptError::Corrupted("Not enough data for Escher record header".to_string()));
        }

        // Read record header (8 bytes) - little-endian format using unsafe for performance
        let record_type = unsafe { u16::from_le_bytes(*(&data[offset..offset + 2] as *const [u8] as *const [u8; 2])) };
        let data_length = unsafe { u32::from_le_bytes(*(&data[offset + 2..offset + 6] as *const [u8] as *const [u8; 4])) };

        // Version and instance are packed in the same 16-bit field
        // Format: VVVV VVVV IIII IIII (V = version bits, I = instance bits)
        let version_instance = unsafe { u16::from_le_bytes(*(&data[offset + 6..offset + 8] as *const [u8] as *const [u8; 2])) };
        let version = (version_instance >> 4) & 0x0FFF;  // High 12 bits for version
        let instance = version_instance & 0x0FFF;        // Low 12 bits for instance

        let record_type_enum = EscherRecordType::from(record_type);
        let total_size = 8 + data_length as usize;

        if offset + total_size > data.len() {
            return Err(PptError::Corrupted("Record extends beyond data bounds".to_string()));
        }

        // Use slice reference where possible to avoid allocation
        let record_data = data[offset + 8..offset + total_size].to_vec();
        let mut record = EscherRecord {
            record_type: record_type_enum,
            version,
            instance,
            data_length,
            data: record_data,
            children: Vec::new(),
            properties: Vec::new(),
        };

        // Pre-allocate children vector if this is a container
        if matches!(record_type_enum, EscherRecordType::Container) && data_length > 0 {
            // Estimate number of children based on data size (rough heuristic)
            let estimated_children = (data_length as usize / 32).min(100); // Assume ~32 bytes per child
            record.children = Vec::with_capacity(estimated_children);
            record.children = Self::parse_container_children(&data[offset + 8..offset + total_size])?;
        }

        // Parse properties if this is an Options record
        if matches!(record_type_enum, EscherRecordType::Options) && data_length > 0 {
            // Options record format: number of properties (2 bytes) + property data
            if record.data.len() >= 2 {
                let num_properties = unsafe { u16::from_le_bytes(*(&record.data[0..2] as *const [u8] as *const [u8; 2])) };
                let property_data = &record.data[2..];

                if let Ok(mut properties) = EscherProperty::parse_properties(property_data, num_properties) {
                    // Pre-allocate with exact size to avoid reallocations
                    record.properties = Vec::with_capacity(properties.len());
                    record.properties.append(&mut properties);
                }
            }
        }

        Ok((record, total_size))
    }

    /// Parse child records from a container record.
    fn parse_container_children(data: &[u8]) -> Result<Vec<EscherRecord>> {
        let mut children = Vec::new();
        let mut offset = 0;

        while offset < data.len() {
            if offset + 8 > data.len() {
                break; // Not enough data for another record header
            }

            let (child, consumed) = Self::parse(data, offset)?;
            children.push(child);
            offset += consumed;
        }

        Ok(children)
    }

    /// Find a child record of a specific type.
    pub fn find_child(&self, record_type: EscherRecordType) -> Option<&EscherRecord> {
        self.children.iter().find(|child| child.record_type == record_type)
    }

    /// Find all child records of a specific type.
    pub fn find_children(&self, record_type: EscherRecordType) -> Vec<&EscherRecord> {
        self.children.iter().filter(|child| child.record_type == record_type).collect()
    }

    /// Find a property by property number.
    pub fn find_property(&self, property_number: u32) -> Option<&EscherProperty> {
        self.properties.iter().find(|prop| prop.property_number() as u32 == property_number)
    }

    /// Get all properties of this record.
    pub fn properties(&self) -> &[EscherProperty] {
        &self.properties
    }

    /// Extract property values for common shape properties.
    /// This provides a convenient interface for accessing frequently used properties.
    pub fn extract_property_values(&self) -> PropertyValues {
        let mut values = PropertyValues::default();

        for property in &self.properties {
            match property.property_number() as u32 {
                // Fill properties
                0x00BF => values.fill_type = Some(property.data as u16),
                0x00C0 => values.fill_color = Some(property.data),
                0x00C1 => values.fill_opacity = Some((property.data & 0xFFFF) as u16),
                0x00C2 => values.fill_back_color = Some(property.data),

                // Line properties
                0x0140 => values.line_color = Some(property.data),
                0x0141 => values.line_opacity = Some((property.data & 0xFFFF) as u16),
                0x0142 => values.line_width = Some(property.data as u16),
                0x0143 => values.line_style = Some(property.data as u16),
                0x0144 => values.line_dash_style = Some(property.data as u16),

                // Shadow properties
                0x0180 => values.shadow_type = Some(property.data as u16),
                0x0181 => values.shadow_color = Some(property.data),
                0x0182 => values.shadow_opacity = Some((property.data & 0xFFFF) as u16),
                0x0183 => values.shadow_offset_x = Some(property.data as i16 as i32),
                0x0184 => values.shadow_offset_y = Some(property.data as i16 as i32),

                // Text properties
                0x01C0 => values.text_left_margin = Some(property.data as i32),
                0x01C1 => values.text_top_margin = Some(property.data as i32),
                0x01C2 => values.text_right_margin = Some(property.data as i32),
                0x01C3 => values.text_bottom_margin = Some(property.data as i32),
                0x01C4 => values.text_anchor = Some(property.data as u16),

                // Transform properties
                0x0000 => values.rotation = Some(property.data),
                0x0001 => values.lock_aspect_ratio = Some(property.data != 0),

                _ => {} // Ignore unknown properties for now
            }
        }

        values
    }

    /// Extract shape properties from this record and its children.
    /// This follows POI's HSLF shape property extraction logic.
    pub fn extract_shape_properties(&self) -> Result<ShapeProperties> {
        let mut props = ShapeProperties::default();

        // Extract transform information (position, size, rotation)
        if let Some(transform) = self.find_child(EscherRecordType::Transform) {
            Self::parse_transform_record(transform, &mut props)?;
        }

        // Extract shape type and ID from shape properties record
        if let Some(shape_props) = self.find_child(EscherRecordType::ShapeProperties) {
            Self::parse_shape_properties_record(shape_props, &mut props)?;
        }

        // Extract additional properties from other records
        Self::extract_additional_properties(self, &mut props)?;

        Ok(props)
    }

    /// Parse transform record data (position, size, rotation).
    /// Based on POI's EscherSpRecord parsing.
    fn parse_transform_record(transform: &EscherRecord, props: &mut ShapeProperties) -> Result<()> {
        // Transform record should have at least 16 bytes for position and size
        if transform.data.len() >= 16 {
            // Parse position (x, y) - 8 bytes each
            props.x = i32::from_le_bytes([
                transform.data[0], transform.data[1], transform.data[2], transform.data[3]
            ]);
            props.y = i32::from_le_bytes([
                transform.data[4], transform.data[5], transform.data[6], transform.data[7]
            ]);

            // Parse size (width, height) - 8 bytes each
            props.width = i32::from_le_bytes([
                transform.data[8], transform.data[9], transform.data[10], transform.data[11]
            ]);
            props.height = i32::from_le_bytes([
                transform.data[12], transform.data[13], transform.data[14], transform.data[15]
            ]);

            // Parse rotation if available (2 bytes)
            if transform.data.len() >= 18 {
                props.rotation = u16::from_le_bytes([transform.data[16], transform.data[17]]);
            }
        }

        Ok(())
    }

    /// Parse shape properties record (type, ID, flags).
    /// Based on POI's EscherSpRecord parsing.
    fn parse_shape_properties_record(shape_props: &EscherRecord, props: &mut ShapeProperties) -> Result<()> {
        if shape_props.data.len() >= 4 { // Shape properties should have at least 4 bytes
            // First 2 bytes: shape type
            let shape_type_id = u16::from_le_bytes([shape_props.data[0], shape_props.data[1]]);
            props.shape_type = ShapeType::from(shape_type_id);

            // Next 2 bytes: shape ID (not 4 bytes as I initially thought)
            if shape_props.data.len() >= 4 {
                props.id = u16::from_le_bytes([shape_props.data[2], shape_props.data[3]]) as u32;
            }

            // Parse flags if available (2 bytes)
            if shape_props.data.len() >= 6 {
                let flags = u16::from_le_bytes([shape_props.data[4], shape_props.data[5]]);
                props.hidden = (flags & 0x0001) != 0; // Hidden flag
            }
        }

        Ok(())
    }

    /// Extract additional properties from various Escher records.
    fn extract_additional_properties(record: &EscherRecord, props: &mut ShapeProperties) -> Result<()> {
        // Check if this record has properties (Options record)
        if !record.properties.is_empty() {
            let prop_values = record.extract_property_values();

            // Apply fill properties
            if let Some(fill_color) = prop_values.fill_color {
                props.fill_color = Some(fill_color);
            }

            // Apply line properties
            if let Some(line_color) = prop_values.line_color {
                props.line_color = Some(line_color);
            }
            if let Some(line_width) = prop_values.line_width {
                props.line_width = Some(line_width);
            }

            // Apply shadow properties
            // Shadow properties would be applied here if available
        }

        // Also check child records for specific property types
        // Parse fill properties
        if let Some(fill_props) = record.find_child(EscherRecordType::FillProperties) {
            Self::parse_fill_properties(fill_props, props)?;
        }

        // Parse line properties
        if let Some(line_props) = record.find_child(EscherRecordType::LineProperties) {
            Self::parse_line_properties(line_props, props)?;
        }

        // Parse shadow properties
        if let Some(shadow_props) = record.find_child(EscherRecordType::ShadowProperties) {
            Self::parse_shadow_properties(shadow_props, props)?;
        }

        Ok(())
    }

    /// Parse fill properties (colors, patterns).
    fn parse_fill_properties(fill_props: &EscherRecord, props: &mut ShapeProperties) -> Result<()> {
        // Fill properties record contains fill-related data
        // For now, extract basic fill information from the record data
        if !fill_props.data.is_empty() {
            // POI's fill parsing logic would go here
            // This is a simplified implementation
            if fill_props.data.len() >= 4 {
                // Extract fill color if available
                let color = u32::from_le_bytes([
                    fill_props.data[0], fill_props.data[1],
                    fill_props.data[2], fill_props.data[3]
                ]);
                props.fill_color = Some(color);
            }
        }
        Ok(())
    }

    /// Parse line properties (color, width, style).
    fn parse_line_properties(line_props: &EscherRecord, props: &mut ShapeProperties) -> Result<()> {
        // Line properties record contains line-related data
        if !line_props.data.is_empty() && line_props.data.len() >= 8 {
            // Extract line color and width
            let color = u32::from_le_bytes([
                line_props.data[0], line_props.data[1],
                line_props.data[2], line_props.data[3]
            ]);
            let width = u16::from_le_bytes([line_props.data[4], line_props.data[5]]);

            props.line_color = Some(color);
            props.line_width = Some(width);
        }
        Ok(())
    }

    /// Parse shadow properties (color, offset, blur).
    fn parse_shadow_properties(_shadow_props: &EscherRecord, _props: &mut ShapeProperties) -> Result<()> {
        // Shadow properties record contains shadow-related data
        // POI would parse this for shadow effects
        // For now, this is a placeholder implementation
        Ok(())
    }

    /// Extract placeholder information from this record.
    /// This follows POI's OEPlaceholderAtom parsing logic.
    ///
    /// OEPlaceholderAtom format (from POI's EscherPlaceholder.fillFields):
    /// - position (4 bytes at offset 8) - placement ID
    /// - placementId (1 byte at offset 12) - placeholder ID
    /// - size (1 byte at offset 13) - placeholder size
    /// - unused (2 bytes at offset 14)
    ///
    /// Returns (placeholder_id, placeholder_size, placement_id)
    pub fn extract_placeholder_info(&self) -> Result<Option<(u16, u8, u16)>> {
        // Look for PlaceholderData record (OEPlaceholderAtom)
        if let Some(placeholder_data) = self.find_child(EscherRecordType::PlaceholderData) {
            // POI's OEPlaceholderAtom structure (8 bytes):
            // Offset 0-3: position/placementId (4 bytes, little-endian) - i32
            // Offset 4: placeholderId (1 byte)
            // Offset 5: size (1 byte)  
            // Offset 6-7: unused (2 bytes)

            if placeholder_data.data.len() >= 8 {
                // Position/placement ID (4 bytes)
                let placement_id = u32::from_le_bytes([
                    placeholder_data.data[0],
                    placeholder_data.data[1],
                    placeholder_data.data[2],
                    placeholder_data.data[3],
                ]) as u16; // Convert to u16 for compatibility

                // Placeholder ID (1 byte at offset 4)
                let placeholder_id = placeholder_data.data[4] as u16;

                // Placeholder size (1 byte at offset 5)
                let placeholder_size = placeholder_data.data[5];

                return Ok(Some((placeholder_id, placeholder_size, placement_id)));
            }
        }

        // Also check if this record itself is a PlaceholderData record
        if self.record_type == EscherRecordType::PlaceholderData && self.data.len() >= 8 {
            let placement_id = u32::from_le_bytes([
                self.data[0],
                self.data[1],
                self.data[2],
                self.data[3],
            ]) as u16;

            let placeholder_id = self.data[4] as u16;
            let placeholder_size = self.data[5];

            return Ok(Some((placeholder_id, placeholder_size, placement_id)));
        }

        Ok(None)
    }

    /// Extract text content from this record.
    /// This follows POI's text extraction logic for Escher text records.
    pub fn extract_text(&self) -> Result<String> {
        if let Some(text_record) = self.find_child(EscherRecordType::Text) {
            Self::parse_text_record(text_record)
        } else {
            Ok(String::new())
        }
    }


    /// Parse text record data according to MS-ODRAW text record format.
    /// Based on POI's EscherTextboxWrapper and related text parsing.
    ///
    /// This properly parses child PPT records (TextCharsAtom, TextBytesAtom, etc.)
    /// from the Escher textbox data.
    fn parse_text_record(text_record: &EscherRecord) -> Result<String> {
        let text_data = &text_record.data;

        if text_data.is_empty() {
            return Ok(String::new());
        }

        // Use EscherTextboxWrapper to properly parse the textbox data
        // This follows POI's approach of wrapping the Escher textbox and
        // extracting child PPT records
        match super::super::escher_textbox::EscherTextboxWrapper::new(text_data.to_vec()) {
            Ok(wrapper) => Ok(wrapper.text().to_string()),
            Err(_) => {
                // Fallback: try simple UTF-16LE decoding for backward compatibility
                if text_data.len() >= 2 {
                    let start_offset = if text_data.len() >= 2 &&
                        text_data[0] == 0xFF && text_data[1] == 0xFE {
                        2 // Skip BOM
                    } else {
                        0
                    };

                    let mut text = String::new();
                    let mut i = start_offset;

                    while i + 1 < text_data.len() {
                        let code_unit = u16::from_le_bytes([text_data[i], text_data[i + 1]]);
                        i += 2;

                        match CharacterAction::process_text_character(code_unit) {
                            CharacterAction::Add(ch) => text.push(ch),
                            CharacterAction::Stop => {
                                // Stop processing, but trim any trailing nulls
                                break;
                            }
                            CharacterAction::Skip => continue,
                        }
                    }

                    Ok(text.trim_end_matches('\u{0}').to_string())
                } else {
                    Ok(String::new())
                }
            }
        }
    }

    /// Parse a complete shape from Escher data.
    pub fn parse_shape(data: &[u8]) -> Result<ShapeProperties> {
        if data.len() < 8 {
            return Err(PptError::Corrupted("Shape data too short".to_string()));
        }

        let (record, _) = Self::parse(data, 0)?;
        record.extract_shape_properties()
    }
}

/// Parser for Escher-based shape data.
pub struct EscherParser {
    /// Parsed records by key (type + instance)
    records: HashMap<u32, EscherRecord>,
    /// Records by shape ID (for placeholder lookup)
    shape_records: HashMap<u32, EscherRecord>,
    /// Placeholder data records (OEPlaceholderAtom)
    placeholder_records: Vec<EscherRecord>,
}

impl EscherParser {
    /// Create a new Escher parser.
    pub fn new() -> Self {
        Self {
            records: HashMap::new(),
            shape_records: HashMap::new(),
            placeholder_records: Vec::new(),
        }
    }

    /// Parse Escher data and extract all records.
    pub fn parse_data(&mut self, data: &[u8]) -> Result<()> {
        let mut offset = 0;

        while offset < data.len() {
            if offset + 8 > data.len() {
                break; // Not enough data for another record
            }

            let (record, consumed) = EscherRecord::parse(data, offset)?;

            // Store record by type/instance for quick lookup
            if record.record_type != EscherRecordType::Container {
                let key = (record.record_type.as_u16() as u32) << 16 | (record.instance as u32);
                self.records.insert(key, record.clone());
            }

            // Also store by shape ID if this record has shape properties
            if record.record_type == EscherRecordType::ShapeProperties && record.data.len() >= 4 {
                let shape_id = u32::from_le_bytes([record.data[2], record.data[3], 0, 0]);
                self.shape_records.insert(shape_id, record.clone());
            }

            // Also store PlaceholderData records for easy access
            if record.record_type == EscherRecordType::PlaceholderData {
                self.placeholder_records.push(record.clone());
            }

            offset += consumed;
        }

        Ok(())
    }

    /// Find a record by type and instance.
    pub fn find_record(&self, record_type: EscherRecordType, instance: u16) -> Option<&EscherRecord> {
        let key = (record_type.as_u16() as u32) << 16 | (instance as u32);
        self.records.get(&key)
    }

    /// Find a record by shape ID.
    pub fn find_record_by_shape_id(&self, shape_id: u32) -> Option<&EscherRecord> {
        self.shape_records.get(&shape_id)
    }

    /// Get all placeholder data records.
    pub fn placeholder_records(&self) -> &[EscherRecord] {
        &self.placeholder_records
    }

    /// Extract all shape properties from parsed data.
    pub fn extract_shapes(&self) -> Result<Vec<ShapeProperties>> {
        let mut shapes = Vec::new();

        for record in self.records.values() {
            if matches!(record.record_type, EscherRecordType::ShapeProperties)
                && let Ok(shape_props) = record.extract_shape_properties() {
                    shapes.push(shape_props);
                }
        }

        Ok(shapes)
    }
}

impl Default for EscherParser {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_escher_record_creation() {
        let record = EscherRecord {
            record_type: EscherRecordType::Transform,
            version: 1,
            instance: 0,
            data_length: 16,
            data: vec![1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16],
            children: Vec::new(),
            properties: Vec::new(),
        };

        assert_eq!(record.record_type, EscherRecordType::Transform);
        assert_eq!(record.version, 1);
        assert_eq!(record.data_length, 16);
        assert_eq!(record.data.len(), 16);
        assert!(record.properties.is_empty());
    }

    #[test]
    fn test_escher_record_type_conversion() {
        assert_eq!(EscherRecordType::from(0xF000), EscherRecordType::Container);
        assert_eq!(EscherRecordType::from(0xF004), EscherRecordType::ShapeProperties);
        assert_eq!(EscherRecordType::from(0xF010), EscherRecordType::Transform);
        assert_eq!(EscherRecordType::from(0xF011), EscherRecordType::Text);
        assert_eq!(EscherRecordType::from(999), EscherRecordType::Container);
    }

    #[test]
    fn test_shape_properties_extraction() {
        // Create a mock transform record
        let transform_record = EscherRecord {
            record_type: EscherRecordType::Transform,
            version: 1,
            instance: 0,
            data_length: 16,
            data: vec![
                0x10, 0x00, 0x00, 0x00, // x = 16
                0x20, 0x00, 0x00, 0x00, // y = 32
                0x64, 0x00, 0x00, 0x00, // width = 100
                0x32, 0x00, 0x00, 0x00, // height = 50
            ],
            children: Vec::new(),
            properties: Vec::new(),
        };

        // Create a mock shape properties record
        let shape_props_record = EscherRecord {
            record_type: EscherRecordType::ShapeProperties,
            version: 1,
            instance: 0,
            data_length: 4,
            data: vec![0x01, 0x00, 0x01, 0x00], // shape_type = 1 (TextBox), id = 1
            children: Vec::new(),
            properties: Vec::new(),
        };

        // Create container record
        let container = EscherRecord {
            record_type: EscherRecordType::Container,
            version: 1,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![transform_record, shape_props_record],
            properties: Vec::new(),
        };

        let props = container.extract_shape_properties().unwrap();
        assert_eq!(props.x, 16);
        assert_eq!(props.y, 32);
        assert_eq!(props.width, 100);
        assert_eq!(props.height, 50);
        assert_eq!(props.shape_type, ShapeType::TextBox);
        assert_eq!(props.id, 1);
    }

    #[test]
    fn test_text_extraction() {
        // Create a container with a text record child
        let text_record = EscherRecord {
            record_type: EscherRecordType::Text,
            version: 1,
            instance: 0,
            data_length: 10,
            data: vec![
                0x48, 0x00, // 'H'
                0x65, 0x00, // 'e'
                0x6C, 0x00, // 'l'
                0x6C, 0x00, // 'l'
                0x6F, 0x00, // 'o'
                0x00, 0x00, // null terminator
            ],
            children: Vec::new(),
            properties: Vec::new(),
        };

        let container = EscherRecord {
            record_type: EscherRecordType::Container,
            version: 1,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![text_record],
            properties: Vec::new(),
        };

        let text = container.extract_text().unwrap();
        // Text record contains "Hello" followed by null terminator
        assert_eq!(text, "Hello");
    }
}
