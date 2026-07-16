use super::super::package::{PptError, Result};
use super::super::records::PptRecord;
use super::super::text_extensions::{
    TextStyleExtension9, TextStyleExtension10, TextStyleExtension11,
};
/// Escher record parsing for PowerPoint shapes.
///
/// This module provides functionality to parse Escher binary records
/// that contain shape data in PowerPoint presentations.
///
/// Escher is Microsoft's binary format for storing graphics and shape data
/// in Office documents, including PowerPoint presentations.
use super::shape::{ShapeProperties, ShapeType};
pub use crate::escher::types::EscherRecordType;
use litchi_core::unit::ppt_master_i64_to_emu_i32;
use std::borrow::Cow;
use std::collections::HashMap;
use zerocopy::{
    FromBytes,
    byteorder::{I32, LittleEndian, U16, U32},
};

/// Escher property types for Office Drawing properties
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum EscherPropertyType {
    /// Transform properties
    Transform,
    /// Shape protection properties
    Protection,
    /// Text box properties
    Text,
    /// WordArt text properties
    GeoText,
    /// Picture properties
    Blip,
    /// Geometry properties
    Geometry,
    /// Fill style properties
    FillStyle,
    /// Line style properties
    LineStyle,
    /// Shadow style properties
    ShadowStyle,
    /// Perspective properties
    Perspective,
    /// 3D properties
    Properties3D,
    /// Shape properties
    Shape,
    /// Callout properties
    Callout,
    /// Group shape properties
    GroupShape,
    /// Unknown property type
    Unknown,
}

impl From<u16> for EscherPropertyType {
    fn from(value: u16) -> Self {
        match value {
            0x0000..=0x003F => EscherPropertyType::Transform,
            0x0040..=0x007F => EscherPropertyType::Protection,
            0x0080..=0x00BF => EscherPropertyType::Text,
            0x00C0..=0x00FF => EscherPropertyType::GeoText,
            0x0100..=0x013F => EscherPropertyType::Blip,
            0x0140..=0x017F => EscherPropertyType::Geometry,
            0x0180..=0x01BF => EscherPropertyType::FillStyle,
            0x01C0..=0x01FF => EscherPropertyType::LineStyle,
            0x0200..=0x023F => EscherPropertyType::ShadowStyle,
            0x0240..=0x027F => EscherPropertyType::Perspective,
            0x0280..=0x02FF => EscherPropertyType::Properties3D,
            0x0300..=0x033F => EscherPropertyType::Shape,
            0x0340..=0x037F => EscherPropertyType::Callout,
            0x0380..=0x03BF => EscherPropertyType::GroupShape,
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

/// An Escher property containing binary data and metadata.
/// Uses `Cow` for zero-copy parsing when possible.
#[derive(Debug, Clone)]
pub struct EscherProperty<'a> {
    /// Property ID (includes type, complex flag, blip flag)
    pub id: u16,
    /// Property data value
    pub data: u32,
    /// Complex data (for complex properties) - uses Cow to avoid unnecessary clones
    pub complex_data: Option<Cow<'a, [u8]>>,
    /// Array data (for array properties) - uses Cow to avoid unnecessary clones
    pub array_data: Option<Cow<'a, [u8]>>,
}

impl<'a> EscherProperty<'a> {
    /// Create a new Escher property
    pub fn new(id: u16, data: u32) -> Self {
        Self {
            id,
            data,
            complex_data: None,
            array_data: None,
        }
    }

    /// Create a complex Escher property with borrowed data (zero-copy)
    pub fn new_complex_borrowed(id: u16, data: u32, complex_data: &'a [u8]) -> Self {
        Self {
            id,
            data,
            complex_data: Some(Cow::Borrowed(complex_data)),
            array_data: None,
        }
    }

    /// Create a complex Escher property with owned data
    pub fn new_complex_owned(id: u16, data: u32, complex_data: Vec<u8>) -> Self {
        Self {
            id,
            data,
            complex_data: Some(Cow::Owned(complex_data)),
            array_data: None,
        }
    }

    /// Create an array Escher property with borrowed data (zero-copy)
    pub fn new_array_borrowed(id: u16, data: u32, array_data: &'a [u8]) -> Self {
        Self {
            id,
            data,
            complex_data: None,
            array_data: Some(Cow::Borrowed(array_data)),
        }
    }

    /// Create an array Escher property with owned data
    pub fn new_array_owned(id: u16, data: u32, array_data: Vec<u8>) -> Self {
        Self {
            id,
            data,
            complex_data: None,
            array_data: Some(Cow::Owned(array_data)),
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
        match self.property_number() {
            0x007F | 0x00BF | 0x00FF | 0x013F | 0x017F | 0x01BF | 0x01FF | 0x023F | 0x027F
            | 0x02BF | 0x033F | 0x057F | 0x05BF | 0x05FF | 0x063F => EscherPropertyHolder::Boolean,
            0x0181 | 0x0183 | 0x01C0 | 0x01C2 | 0x0201 | 0x0287 | 0x02BE => {
                EscherPropertyHolder::RGB
            },
            0x0144 => EscherPropertyHolder::ShapePath,
            0x0145 | 0x0146 | 0x0197 | 0x01CF | 0x0383 | 0x03A0 => EscherPropertyHolder::Array,
            _ if self.is_complex() => EscherPropertyHolder::Complex,
            _ => EscherPropertyHolder::Simple,
        }
    }

    /// Parse properties from binary data (based on POI's EscherPropertyFactory).
    /// Optimized for performance with pre-allocation and zero-copy when possible.
    pub fn parse_properties(data: &'a [u8], num_properties: u16) -> Result<Vec<Self>> {
        if num_properties == 0 {
            return Ok(Vec::new());
        }

        let property_count = usize::from(num_properties);
        let header_size = property_count.checked_mul(6).ok_or_else(|| {
            PptError::Corrupted("Escher property header size overflow".to_string())
        })?;
        if header_size > data.len() {
            return Err(PptError::Corrupted(format!(
                "Escher property headers require {header_size} bytes, found {}",
                data.len()
            )));
        }

        let mut descriptors = Vec::with_capacity(property_count);
        for index in 0..property_count {
            let offset = index * 6;
            let prop_id = U16::<LittleEndian>::read_from_bytes(&data[offset..offset + 2])
                .map(|v| v.get())
                .unwrap_or(0);
            let prop_data = U32::<LittleEndian>::read_from_bytes(&data[offset + 2..offset + 6])
                .map(|v| v.get())
                .unwrap_or(0);
            descriptors.push((prop_id, prop_data));
        }

        let mut properties = Vec::with_capacity(property_count);
        let mut complex_offset = header_size;
        for (prop_id, prop_data) in descriptors {
            let mut property = Self::new(prop_id, prop_data);
            if property.is_complex() {
                let declared_size = usize::try_from(prop_data).map_err(|_| {
                    PptError::Corrupted("Escher complex property size overflow".to_string())
                })?;
                let complex_size = if property.property_holder() == EscherPropertyHolder::Array
                    && declared_size > 0
                    && complex_offset
                        .checked_add(6)
                        .is_some_and(|end| end <= data.len())
                {
                    let count = usize::from(u16::from_le_bytes([
                        data[complex_offset],
                        data[complex_offset + 1],
                    ]));
                    let raw_element_size =
                        i16::from_le_bytes([data[complex_offset + 4], data[complex_offset + 5]]);
                    let element_size = if raw_element_size < 0 {
                        usize::from(raw_element_size.unsigned_abs() >> 2)
                    } else {
                        raw_element_size as usize
                    };
                    let payload_size = count.checked_mul(element_size).ok_or_else(|| {
                        PptError::Corrupted("Escher array property size overflow".to_string())
                    })?;
                    if payload_size == declared_size {
                        declared_size.checked_add(6).ok_or_else(|| {
                            PptError::Corrupted("Escher array property size overflow".to_string())
                        })?
                    } else {
                        declared_size
                    }
                } else {
                    declared_size
                };
                let complex_end = complex_offset.checked_add(complex_size).ok_or_else(|| {
                    PptError::Corrupted("Escher complex property offset overflow".to_string())
                })?;
                if complex_end > data.len() {
                    return Err(PptError::Corrupted(format!(
                        "Escher complex property requires {complex_size} bytes, found {}",
                        data.len().saturating_sub(complex_offset)
                    )));
                }
                let complex_data = &data[complex_offset..complex_end];
                property = if property.property_holder() == EscherPropertyHolder::Array {
                    Self::new_array_borrowed(prop_id, prop_data, complex_data)
                } else {
                    Self::new_complex_borrowed(prop_id, prop_data, complex_data)
                };
                complex_offset = complex_end;
            }
            properties.push(property);
        }
        Ok(properties)
    }
}

/// Property values extracted from Escher records for convenient access
#[derive(Debug, Clone, Default)]
pub struct PropertyValues {
    // Fill properties
    pub fill_type: Option<u32>,
    pub fill_color: Option<u32>,
    pub fill_opacity: Option<u32>,
    pub fill_back_color: Option<u32>,

    // Line properties
    pub line_color: Option<u32>,
    pub line_opacity: Option<u32>,
    pub line_width: Option<i32>,
    pub line_style: Option<u32>,
    pub line_dash_style: Option<u32>,

    // Shadow properties
    pub shadow_type: Option<u32>,
    pub shadow_color: Option<u32>,
    pub shadow_opacity: Option<u32>,
    pub shadow_offset_x: Option<i32>,
    pub shadow_offset_y: Option<i32>,
    pub shadow_enabled: Option<bool>,
    pub shadow_obscured: Option<bool>,

    // Text properties
    pub text_left_margin: Option<i32>,
    pub text_top_margin: Option<i32>,
    pub text_right_margin: Option<i32>,
    pub text_bottom_margin: Option<i32>,
    pub text_anchor: Option<u32>,

    // Transform properties
    pub rotation: Option<i32>,
    pub lock_aspect_ratio: Option<bool>,
}

/// An Escher record containing binary data and metadata.
/// Optimized for performance with zero-copy parsing using `Cow`.
#[derive(Debug, Clone)]
pub struct EscherRecord<'a> {
    /// Record type
    pub record_type: EscherRecordType,
    /// Record version
    pub version: u16,
    /// Record instance (sub-type)
    pub instance: u16,
    /// Record data length
    pub data_length: u32,
    /// Record data - uses Cow to avoid unnecessary clones during parsing
    pub data: Cow<'a, [u8]>,
    /// Child records (for container records)
    pub children: Vec<EscherRecord<'a>>,
    /// Parsed properties (for Options records)
    pub properties: Vec<EscherProperty<'a>>,
}

impl<'a> EscherRecord<'a> {
    /// Parse an Escher record from binary data with zero-copy optimization.
    /// Uses `Cow` to borrow data when possible, avoiding unnecessary allocations.
    ///
    /// # Arguments
    ///
    /// * `data` - Binary data containing the record
    /// * `offset` - Starting offset in the data
    ///
    /// # Returns
    ///
    /// Tuple of (parsed_record, bytes_consumed)
    pub fn parse(data: &'a [u8], offset: usize) -> Result<(Self, usize)> {
        let header_end = offset.checked_add(8).ok_or_else(|| {
            PptError::Corrupted("Escher record header offset overflow".to_string())
        })?;
        if header_end > data.len() {
            return Err(PptError::Corrupted(
                "Not enough data for Escher record header".to_string(),
            ));
        }

        // OfficeArtRecordHeader: recVer/recInstance, recType, then recLen.
        let version_instance = U16::<LittleEndian>::read_from_bytes(&data[offset..offset + 2])
            .map(|v| v.get())
            .unwrap_or(0);
        let record_type = U16::<LittleEndian>::read_from_bytes(&data[offset + 2..offset + 4])
            .map(|v| v.get())
            .unwrap_or(0);
        let data_length = U32::<LittleEndian>::read_from_bytes(&data[offset + 4..offset + 8])
            .map(|v| v.get())
            .unwrap_or(0);
        let version = version_instance & 0x000F;
        let instance = version_instance >> 4;

        let record_type_enum = EscherRecordType::from(record_type);
        let total_size = 8usize
            .checked_add(data_length as usize)
            .ok_or_else(|| PptError::Corrupted("Escher record size overflow".to_string()))?;
        let record_end = offset
            .checked_add(total_size)
            .ok_or_else(|| PptError::Corrupted("Escher record offset overflow".to_string()))?;
        if record_end > data.len() {
            return Err(PptError::Corrupted(
                "Record extends beyond data bounds".to_string(),
            ));
        }

        // Use zero-copy borrowing for record data
        let payload = &data[header_end..record_end];
        let mut record = EscherRecord {
            record_type: record_type_enum,
            version,
            instance,
            data_length,
            data: Cow::Borrowed(payload),
            children: Vec::new(),
            properties: Vec::new(),
        };

        if version == 0x000F && !payload.is_empty() {
            record.children = Self::parse_container_children(payload)?;
        }

        if matches!(
            record_type_enum,
            EscherRecordType::Opt | EscherRecordType::SecondaryOpt | EscherRecordType::TertiaryOpt
        ) {
            record.properties = EscherProperty::parse_properties(payload, instance)?;
        }

        Ok((record, total_size))
    }

    /// Parse child records from a container record with zero-copy optimization.
    fn parse_container_children(data: &'a [u8]) -> Result<Vec<EscherRecord<'a>>> {
        let mut children = Vec::new();
        let mut offset = 0;

        while offset < data.len() {
            if offset + 8 > data.len() {
                return Err(PptError::Corrupted(
                    "Truncated child Escher record header".to_string(),
                ));
            }

            let (child, consumed) = Self::parse(data, offset)?;
            children.push(child);
            offset += consumed;
        }

        Ok(children)
    }

    /// Find a child record of a specific type.
    pub fn find_child(&self, record_type: EscherRecordType) -> Option<&EscherRecord<'a>> {
        self.children
            .iter()
            .find(|child| child.record_type == record_type)
    }

    /// Find all child records of a specific type.
    pub fn find_children(&self, record_type: EscherRecordType) -> Vec<&EscherRecord<'a>> {
        self.children
            .iter()
            .filter(|child| child.record_type == record_type)
            .collect()
    }

    /// Find a property by property number.
    pub fn find_property(&self, property_number: u32) -> Option<&EscherProperty<'a>> {
        self.properties
            .iter()
            .find(|prop| prop.property_number() as u32 == property_number)
    }

    /// Get all properties of this record.
    pub fn properties(&self) -> &[EscherProperty<'a>] {
        &self.properties
    }

    /// Extract property values for common shape properties.
    /// This provides a convenient interface for accessing frequently used properties.
    pub fn extract_property_values(&self) -> PropertyValues {
        let mut values = PropertyValues::default();

        for property in &self.properties {
            match property.property_number() as u32 {
                // Fill properties
                0x0180 => values.fill_type = Some(property.data),
                0x0181 => values.fill_color = Some(property.data),
                0x0182 => values.fill_opacity = Some(property.data),
                0x0183 => values.fill_back_color = Some(property.data),

                // Line properties
                0x01C0 => values.line_color = Some(property.data),
                0x01C1 => values.line_opacity = Some(property.data),
                0x01CB => values.line_width = Some(property.data as i32),
                0x01CD => values.line_style = Some(property.data),
                0x01CE => values.line_dash_style = Some(property.data),

                // Shadow properties
                0x0200 => values.shadow_type = Some(property.data),
                0x0201 => values.shadow_color = Some(property.data),
                0x0204 => values.shadow_opacity = Some(property.data),
                0x0205 => values.shadow_offset_x = Some(property.data as i32),
                0x0206 => values.shadow_offset_y = Some(property.data as i32),
                0x023F => {
                    let use_bits = property.data >> 16;
                    if use_bits & 0x0002 != 0 {
                        values.shadow_enabled = Some(property.data & 0x0002 != 0);
                    }
                    if use_bits & 0x0001 != 0 {
                        values.shadow_obscured = Some(property.data & 0x0001 != 0);
                    }
                },

                // Text properties
                0x0081 => values.text_left_margin = Some(property.data as i32),
                0x0082 => values.text_top_margin = Some(property.data as i32),
                0x0083 => values.text_right_margin = Some(property.data as i32),
                0x0084 => values.text_bottom_margin = Some(property.data as i32),
                0x0087 => values.text_anchor = Some(property.data),

                // Transform properties
                0x0004 => values.rotation = Some(property.data as i32),
                0x007F => {
                    let use_mask = 0x0080_0000;
                    if property.data & use_mask != 0 {
                        values.lock_aspect_ratio = Some(property.data & 0x0000_0080 != 0);
                    }
                },

                _ => {}, // Ignore unknown properties for now
            }
        }

        values
    }

    /// Extract shape properties from this record and its children.
    /// This follows POI's HSLF shape property extraction logic.
    pub fn extract_shape_properties(&self) -> Result<ShapeProperties> {
        let mut props = ShapeProperties::default();

        if let Some(anchor) = self
            .find_child(EscherRecordType::ClientAnchor)
            .or_else(|| self.find_child(EscherRecordType::ChildAnchor))
        {
            Self::parse_anchor_record(anchor, &mut props)?;
        }

        if let Some(shape_props) = self.find_child(EscherRecordType::Sp) {
            Self::parse_shape_properties_record(shape_props, &mut props)?;
        }

        // Extract additional properties from other records
        Self::extract_additional_properties(self, &mut props)?;

        Ok(props)
    }

    /// Parse a PowerPoint client anchor or an OfficeArt child anchor.
    fn parse_anchor_record(anchor: &EscherRecord, props: &mut ShapeProperties) -> Result<()> {
        let (left, top, right, bottom) = if anchor.record_type == EscherRecordType::ClientAnchor
            && anchor.data.len() >= 8
        {
            let top = i16::from_le_bytes([anchor.data[0], anchor.data[1]]);
            let left = i16::from_le_bytes([anchor.data[2], anchor.data[3]]);
            let right = i16::from_le_bytes([anchor.data[4], anchor.data[5]]);
            let bottom = i16::from_le_bytes([anchor.data[6], anchor.data[7]]);
            (
                ppt_master_i64_to_emu_i32(i64::from(left)),
                ppt_master_i64_to_emu_i32(i64::from(top)),
                ppt_master_i64_to_emu_i32(i64::from(right)),
                ppt_master_i64_to_emu_i32(i64::from(bottom)),
            )
        } else if anchor.record_type == EscherRecordType::ChildAnchor && anchor.data.len() >= 16 {
            (
                I32::<LittleEndian>::read_from_bytes(&anchor.data[0..4])
                    .map(|v| v.get())
                    .unwrap_or(0),
                I32::<LittleEndian>::read_from_bytes(&anchor.data[4..8])
                    .map(|v| v.get())
                    .unwrap_or(0),
                I32::<LittleEndian>::read_from_bytes(&anchor.data[8..12])
                    .map(|v| v.get())
                    .unwrap_or(0),
                I32::<LittleEndian>::read_from_bytes(&anchor.data[12..16])
                    .map(|v| v.get())
                    .unwrap_or(0),
            )
        } else {
            return Err(PptError::Corrupted(
                "Invalid OfficeArt shape anchor".to_string(),
            ));
        };

        props.x = left;
        props.y = top;
        props.width = right.saturating_sub(left);
        props.height = bottom.saturating_sub(top);
        Ok(())
    }

    /// Parse an OfficeArt `Sp` atom.
    fn parse_shape_properties_record(
        shape_props: &EscherRecord,
        props: &mut ShapeProperties,
    ) -> Result<()> {
        if shape_props.data.len() < 8 {
            return Err(PptError::Corrupted(
                "OfficeArt Sp atom is shorter than 8 bytes".to_string(),
            ));
        }

        props.shape_type =
            ShapeType::from(crate::consts::EscherShapeType::from(shape_props.instance));
        props.id = U32::<LittleEndian>::read_from_bytes(&shape_props.data[0..4])
            .map(|v| v.get())
            .unwrap_or(0);
        Ok(())
    }

    /// Extract additional properties from various Escher records.
    fn extract_additional_properties(
        record: &EscherRecord,
        props: &mut ShapeProperties,
    ) -> Result<()> {
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
                let rounded_points = line_width.max(0).saturating_add(6_350) / 12_700;
                props.line_width = Some(u16::try_from(rounded_points).unwrap_or(u16::MAX));
            }
            if let Some(rotation) = prop_values.rotation {
                let degrees = rotation / 65_536;
                props.rotation = degrees.rem_euclid(360) as u16;
            }
        }

        for child in &record.children {
            Self::extract_additional_properties(child, props)?;
        }

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
        let Some(client_data) = (self.record_type == EscherRecordType::ClientData)
            .then_some(self)
            .or_else(|| self.find_child(EscherRecordType::ClientData))
        else {
            return Ok(None);
        };

        let mut offset = 0usize;
        while offset < client_data.data.len() {
            let header_end = offset.checked_add(8).ok_or_else(|| {
                PptError::Corrupted("Placeholder record offset overflow".to_string())
            })?;
            if header_end > client_data.data.len() {
                return Err(PptError::Corrupted(
                    "Truncated placeholder record header".to_string(),
                ));
            }
            let record_type =
                u16::from_le_bytes([client_data.data[offset + 2], client_data.data[offset + 3]]);
            let length = u32::from_le_bytes([
                client_data.data[offset + 4],
                client_data.data[offset + 5],
                client_data.data[offset + 6],
                client_data.data[offset + 7],
            ]) as usize;
            let record_end = header_end.checked_add(length).ok_or_else(|| {
                PptError::Corrupted("Placeholder record size overflow".to_string())
            })?;
            if record_end > client_data.data.len() {
                return Err(PptError::Corrupted(
                    "Placeholder record extends beyond client data".to_string(),
                ));
            }

            if record_type == 0x0BC3 {
                let payload = &client_data.data[header_end..record_end];
                if payload.len() < 8 {
                    return Err(PptError::Corrupted(
                        "OEPlaceholderAtom is shorter than 8 bytes".to_string(),
                    ));
                }
                let placement_id =
                    u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]) as u16;
                return Ok(Some((payload[4] as u16, payload[5], placement_id)));
            }
            offset = record_end;
        }
        Ok(None)
    }

    /// Extract text content from this record.
    /// This follows POI's text extraction logic for Escher text records.
    pub fn extract_text(&self) -> Result<String> {
        if let Some(text_record) = self.find_child(EscherRecordType::ClientTextbox) {
            Self::parse_text_record(text_record)
        } else {
            Ok(String::new())
        }
    }

    /// Extract the PowerPoint 9 text-style extension stored beside a textbox.
    ///
    /// MS-PPT stores `StyleTextProp9Atom` in the shape's `ClientData`, under
    /// the `___PPT9` programmable binary tag, rather than in `ClientTextbox`.
    pub fn extract_text_style_extension9(&self) -> Result<Option<TextStyleExtension9>> {
        self.extract_versioned_text_style_record(
            9,
            crate::consts::PptRecordType::StyleTextProp9Atom,
        )?
        .map(|record| TextStyleExtension9::parse(&record.data))
        .transpose()
    }

    /// Extract PowerPoint 10 alternate-script font formatting for this text.
    pub fn extract_text_style_extension10(&self) -> Result<Option<TextStyleExtension10>> {
        self.extract_versioned_text_style_record(
            10,
            crate::consts::PptRecordType::StyleTextProp10Atom,
        )?
        .map(|record| TextStyleExtension10::parse(&record.data))
        .transpose()
    }

    /// Extract PowerPoint 11 smart-tag formatting for this text.
    pub fn extract_text_style_extension11(&self) -> Result<Option<TextStyleExtension11>> {
        self.extract_versioned_text_style_record(
            11,
            crate::consts::PptRecordType::StyleTextProp11Atom,
        )?
        .map(|record| TextStyleExtension11::parse(&record.data))
        .transpose()
    }

    fn extract_versioned_text_style_record(
        &self,
        version: u8,
        record_type: crate::consts::PptRecordType,
    ) -> Result<Option<PptRecord>> {
        let Some(client_data) = (self.record_type == EscherRecordType::ClientData)
            .then_some(self)
            .or_else(|| self.find_child(EscherRecordType::ClientData))
        else {
            return Ok(None);
        };

        let mut result = None;
        for record in parse_ppt_record_sequence(&client_data.data, "shape ClientData")? {
            if record.record_type != crate::consts::PptRecordType::ProgTags {
                continue;
            }
            for tag in parse_ppt_record_sequence(&record.data, &format!("PPT{version} ProgTags"))? {
                if tag.record_type != crate::consts::PptRecordType::ProgBinaryTag {
                    continue;
                }
                let tag_children =
                    parse_ppt_record_sequence(&tag.data, &format!("PPT{version} ProgBinaryTag"))?;
                let Some(name) = tag_children
                    .iter()
                    .find(|child| child.record_type == crate::consts::PptRecordType::CString)
                else {
                    continue;
                };
                if !is_versioned_ppt_tag_name(name, version) {
                    continue;
                }
                let blob = tag_children
                    .iter()
                    .find(|child| child.record_type == crate::consts::PptRecordType::BinaryTagData)
                    .ok_or_else(|| {
                        PptError::Corrupted(format!(
                            "___PPT{version} programmable tag is missing BinaryTagData"
                        ))
                    })?;
                for extension_record in parse_ppt_record_sequence(
                    &blob.data,
                    &format!("___PPT{version} BinaryTagData"),
                )? {
                    if extension_record.record_type != record_type {
                        continue;
                    }
                    if extension_record.version != 0 || extension_record.instance != 0 {
                        return Err(PptError::Corrupted(format!(
                            "Versioned text style atom {record_type:?} has an invalid header"
                        )));
                    }
                    if result.is_some() {
                        return Err(PptError::Corrupted(format!(
                            "Shape contains multiple {record_type:?} records"
                        )));
                    }
                    result = Some(extension_record);
                }
            }
        }
        Ok(result)
    }

    /// Parse text record data according to MS-ODRAW text record format.
    /// Based on POI's EscherTextboxWrapper and related text parsing.
    ///
    /// This properly parses child PPT records (TextCharsAtom, TextBytesAtom, etc.)
    /// from the Escher textbox data.
    fn parse_text_record(text_record: &EscherRecord<'a>) -> Result<String> {
        let text_data = &text_record.data;

        if text_data.is_empty() {
            return Ok(String::new());
        }

        let wrapper = super::super::escher_textbox::EscherTextboxWrapper::new(text_data.to_vec())?;
        Ok(wrapper.text().to_string())
    }

    /// Parse a complete shape from Escher data with zero-copy optimization.
    pub fn parse_shape(data: &'a [u8]) -> Result<ShapeProperties> {
        if data.len() < 8 {
            return Err(PptError::Corrupted("Shape data too short".to_string()));
        }

        let (record, _) = Self::parse(data, 0)?;
        record.extract_shape_properties()
    }
}

fn parse_ppt_record_sequence(data: &[u8], context: &str) -> Result<Vec<PptRecord>> {
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let header_end = offset.checked_add(8).ok_or_else(|| {
            PptError::Corrupted(format!("{context} record header offset overflow"))
        })?;
        if header_end > data.len() {
            return Err(PptError::Corrupted(format!(
                "Truncated record header in {context}"
            )));
        }
        let length = u32::from_le_bytes([
            data[offset + 4],
            data[offset + 5],
            data[offset + 6],
            data[offset + 7],
        ]);
        let length = usize::try_from(length)
            .map_err(|_| PptError::Corrupted(format!("{context} record size overflow")))?;
        let record_end = header_end
            .checked_add(length)
            .ok_or_else(|| PptError::Corrupted(format!("{context} record size overflow")))?;
        if record_end > data.len() {
            return Err(PptError::Corrupted(format!(
                "Record extends beyond {context}"
            )));
        }
        let (record, consumed) = PptRecord::parse(&data[offset..record_end], 0)?;
        if consumed != record_end - offset {
            return Err(PptError::Corrupted(format!(
                "Record in {context} was only partially parsed"
            )));
        }
        records.push(record);
        offset = record_end;
    }
    Ok(records)
}

fn is_versioned_ppt_tag_name(record: &PptRecord, version: u8) -> bool {
    let expected: &[u16] = match version {
        9 => &[0x5f, 0x5f, 0x5f, 0x50, 0x50, 0x54, 0x39],
        10 => &[0x5f, 0x5f, 0x5f, 0x50, 0x50, 0x54, 0x31, 0x30],
        11 => &[0x5f, 0x5f, 0x5f, 0x50, 0x50, 0x54, 0x31, 0x31],
        _ => return false,
    };
    record.version == 0
        && record.instance == 0
        && record.data.len() == expected.len() * 2
        && record
            .data
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .eq(expected.iter().copied())
}

/// Parser for Escher-based shape data with zero-copy optimization.
/// Uses `Cow` to avoid unnecessary memory clones during parsing.
pub struct EscherParser<'a> {
    /// Parsed records by key (type + instance)
    records: HashMap<u32, EscherRecord<'a>>,
    /// Records by shape ID (for placeholder lookup)
    shape_records: HashMap<u32, EscherRecord<'a>>,
    /// Placeholder data records (OEPlaceholderAtom)
    placeholder_records: Vec<EscherRecord<'a>>,
}

impl<'a> EscherParser<'a> {
    /// Create a new Escher parser.
    pub fn new() -> Self {
        Self {
            records: HashMap::new(),
            shape_records: HashMap::new(),
            placeholder_records: Vec::new(),
        }
    }

    /// Parse Escher data and extract all records with zero-copy optimization.
    pub fn parse_data(&mut self, data: &'a [u8]) -> Result<()> {
        let mut offset = 0;

        while offset < data.len() {
            if offset + 8 > data.len() {
                break; // Not enough data for another record
            }

            let (record, consumed) = EscherRecord::parse(data, offset)?;
            self.index_record(&record);
            offset += consumed;
        }

        Ok(())
    }

    fn index_record(&mut self, record: &EscherRecord<'a>) {
        if record.version != 0x000F {
            let key = (record.record_type as u16 as u32) << 16 | u32::from(record.instance);
            self.records.insert(key, record.clone());
        }

        if record.record_type == EscherRecordType::SpContainer
            && let Some(shape) = record.find_child(EscherRecordType::Sp)
            && shape.data.len() >= 4
        {
            let shape_id = U32::<LittleEndian>::read_from_bytes(&shape.data[0..4])
                .map(|v| v.get())
                .unwrap_or(0);
            self.shape_records.insert(shape_id, record.clone());
        }

        if record.record_type == EscherRecordType::ClientData {
            self.placeholder_records.push(record.clone());
        }

        for child in &record.children {
            self.index_record(child);
        }
    }

    /// Find a record by type and instance.
    pub fn find_record(
        &self,
        record_type: EscherRecordType,
        instance: u16,
    ) -> Option<&EscherRecord<'a>> {
        let key = (record_type as u16 as u32) << 16 | u32::from(instance);
        self.records.get(&key)
    }

    /// Find a record by shape ID.
    pub fn find_record_by_shape_id(&self, shape_id: u32) -> Option<&EscherRecord<'a>> {
        self.shape_records.get(&shape_id)
    }

    /// Get all placeholder data records.
    pub fn placeholder_records(&self) -> &[EscherRecord<'a>] {
        &self.placeholder_records
    }

    /// Extract all shape properties from parsed data.
    pub fn extract_shapes(&self) -> Result<Vec<ShapeProperties>> {
        self.shape_records
            .values()
            .map(EscherRecord::extract_shape_properties)
            .collect()
    }
}

impl<'a> Default for EscherParser<'a> {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn officeart_record(
        version: u16,
        instance: u16,
        record_type: EscherRecordType,
        payload: &[u8],
    ) -> Vec<u8> {
        let mut data = Vec::with_capacity(8 + payload.len());
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&(record_type as u16).to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn ppt_record(version: u16, instance: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::with_capacity(8 + payload.len());
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&record_type.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn shape_with_versioned_extension(
        tag_name: &str,
        record_type: u16,
        style_version: u16,
        payload: &[u8],
    ) -> EscherRecord<'static> {
        let style = ppt_record(style_version, 0, record_type, payload);
        let blob = ppt_record(0, 0, 0x138b, &style);
        let tag_name: Vec<u8> = tag_name.encode_utf16().flat_map(u16::to_le_bytes).collect();
        let mut tag_payload = ppt_record(0, 0, 4026, &tag_name);
        tag_payload.extend_from_slice(&blob);
        let tag = ppt_record(0x0f, 0, 0x138a, &tag_payload);
        let prog_tags = ppt_record(0x0f, 0, 0x1388, &tag);
        let client_data = EscherRecord {
            record_type: EscherRecordType::ClientData,
            version: 0x0f,
            instance: 0,
            data_length: prog_tags.len() as u32,
            data: Cow::Owned(prog_tags),
            children: Vec::new(),
            properties: Vec::new(),
        };
        EscherRecord {
            record_type: EscherRecordType::SpContainer,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Cow::Borrowed(&[]),
            children: vec![client_data],
            properties: Vec::new(),
        }
    }

    #[test]
    fn extracts_style_text_prop9_from_shape_client_data() {
        let shape = shape_with_versioned_extension("___PPT9", 4012, 0, &[0; 12]);
        let extension = shape.extract_text_style_extension9().unwrap().unwrap();
        assert_eq!(extension.runs.len(), 1);
        assert_eq!(extension.runs[0].paragraph.mask, 0);

        let malformed = shape_with_versioned_extension("___PPT9", 4012, 1, &[0; 12]);
        assert!(malformed.extract_text_style_extension9().is_err());
    }

    #[test]
    fn extracts_powerpoint_10_and_11_text_styles_from_client_data() {
        let mut font_payload = Vec::new();
        font_payload.extend_from_slice(&0x0300_0000u32.to_le_bytes());
        font_payload.extend_from_slice(&17u16.to_le_bytes());
        font_payload.extend_from_slice(&23u16.to_le_bytes());
        let shape = shape_with_versioned_extension("___PPT10", 4017, 0, &font_payload);
        let extension = shape.extract_text_style_extension10().unwrap().unwrap();
        assert_eq!(extension.runs[0].new_east_asian_font_ref, Some(17));
        assert_eq!(extension.runs[0].complex_script_font_ref, Some(23));
        assert!(shape.extract_text_style_extension11().unwrap().is_none());

        let mut smart_tag_payload = Vec::new();
        smart_tag_payload.extend_from_slice(&0x0200u32.to_le_bytes());
        smart_tag_payload.extend_from_slice(&1u32.to_le_bytes());
        smart_tag_payload.extend_from_slice(&99u32.to_le_bytes());
        let shape = shape_with_versioned_extension("___PPT11", 4022, 0, &smart_tag_payload);
        let extension = shape.extract_text_style_extension11().unwrap().unwrap();
        assert_eq!(extension.runs[0].smart_tag_indices, vec![99]);
    }

    #[test]
    fn parses_officeart_headers_options_and_nested_shape_containers() {
        let mut opt_payload = Vec::new();
        opt_payload.extend_from_slice(&0x0181_u16.to_le_bytes());
        opt_payload.extend_from_slice(&0x0012_3456_u32.to_le_bytes());
        let opt = officeart_record(3, 1, EscherRecordType::Opt, &opt_payload);

        let mut shape_payload = Vec::new();
        shape_payload.extend_from_slice(&42_u32.to_le_bytes());
        shape_payload.extend_from_slice(&0x0000_0A00_u32.to_le_bytes());
        let shape = officeart_record(2, 1, EscherRecordType::Sp, &shape_payload);
        let anchor = officeart_record(
            0,
            0,
            EscherRecordType::ClientAnchor,
            &[32, 0, 16, 0, 116, 0, 82, 0],
        );

        let mut shape_container_payload = Vec::new();
        shape_container_payload.extend_from_slice(&shape);
        shape_container_payload.extend_from_slice(&opt);
        shape_container_payload.extend_from_slice(&anchor);
        let shape_container = officeart_record(
            0xF,
            0,
            EscherRecordType::SpContainer,
            &shape_container_payload,
        );
        let drawing = officeart_record(0xF, 0, EscherRecordType::DgContainer, &shape_container);

        let (record, consumed) = EscherRecord::parse(&drawing, 0).unwrap();
        assert_eq!(consumed, drawing.len());
        assert_eq!(record.version, 0xF);
        assert_eq!(record.record_type, EscherRecordType::DgContainer);
        assert_eq!(record.children.len(), 1);
        let parsed_shape = &record.children[0];
        assert_eq!(parsed_shape.record_type, EscherRecordType::SpContainer);
        assert_eq!(parsed_shape.children[1].record_type, EscherRecordType::Opt);
        assert_eq!(parsed_shape.children[1].instance, 1);
        assert_eq!(parsed_shape.children[1].properties.len(), 1);

        let mut parser = EscherParser::new();
        parser.parse_data(&drawing).unwrap();
        let indexed_shape = parser.find_record_by_shape_id(42).unwrap();
        assert_eq!(indexed_shape.record_type, EscherRecordType::SpContainer);
        let properties = indexed_shape.extract_shape_properties().unwrap();
        assert_eq!(properties.id, 42);
        assert_eq!(properties.shape_type, ShapeType::AutoShape);
        assert_eq!(properties.fill_color, Some(0x0012_3456));
        assert_eq!(properties.x, ppt_master_i64_to_emu_i32(16));
        assert_eq!(properties.y, ppt_master_i64_to_emu_i32(32));
        assert_eq!(properties.width, ppt_master_i64_to_emu_i32(100));
        assert_eq!(properties.height, ppt_master_i64_to_emu_i32(50));
        assert_eq!(parser.extract_shapes().unwrap().len(), 1);
    }

    #[test]
    fn classifies_officeart_properties_by_spec_table() {
        assert_eq!(
            EscherProperty::new(0x0004, 0).property_type(),
            EscherPropertyType::Transform
        );
        assert_eq!(
            EscherProperty::new(0x007F, 0).property_type(),
            EscherPropertyType::Protection
        );
        assert_eq!(
            EscherProperty::new(0x0081, 0).property_type(),
            EscherPropertyType::Text
        );
        assert_eq!(
            EscherProperty::new(0x00C0, 0).property_type(),
            EscherPropertyType::GeoText
        );
        assert_eq!(
            EscherProperty::new(0x4104, 0).property_type(),
            EscherPropertyType::Blip
        );
        assert_eq!(
            EscherProperty::new(0x8145, 0).property_holder(),
            EscherPropertyHolder::Array
        );
        assert_eq!(
            EscherProperty::new(0x0181, 0).property_holder(),
            EscherPropertyHolder::RGB
        );
        assert_eq!(
            EscherProperty::new(0x01FF, 0).property_holder(),
            EscherPropertyHolder::Boolean
        );
        assert_eq!(
            EscherProperty::new(0x0144, 0).property_holder(),
            EscherPropertyHolder::ShapePath
        );
        assert_eq!(
            EscherProperty::new(0x8400, 0).property_holder(),
            EscherPropertyHolder::Complex
        );
        assert_eq!(
            EscherProperty::new(0x0400, 0).property_type(),
            EscherPropertyType::Unknown
        );
    }

    #[test]
    fn parses_all_property_headers_before_complex_payloads() {
        let array_data = [2, 0, 2, 0, 2, 0, 1, 2, 3, 4];
        let text_data = [b'A', 0, 0, 0];
        let mut data = Vec::new();
        data.extend_from_slice(&0x8145_u16.to_le_bytes());
        data.extend_from_slice(&4_u32.to_le_bytes()); // Array size excluding its 6-byte header.
        data.extend_from_slice(&0x0181_u16.to_le_bytes());
        data.extend_from_slice(&0x0012_3456_u32.to_le_bytes());
        data.extend_from_slice(&0x80C0_u16.to_le_bytes());
        data.extend_from_slice(&(text_data.len() as u32).to_le_bytes());
        data.extend_from_slice(&array_data);
        data.extend_from_slice(&text_data);

        let properties = EscherProperty::parse_properties(&data, 3).unwrap();
        assert_eq!(properties.len(), 3);
        assert_eq!(
            properties[0].array_data.as_deref(),
            Some(array_data.as_slice())
        );
        assert_eq!(properties[1].data, 0x0012_3456);
        assert_eq!(
            properties[2].complex_data.as_deref(),
            Some(text_data.as_slice())
        );

        let truncated = &data[..18];
        assert!(EscherProperty::parse_properties(truncated, 3).is_err());
    }

    #[test]
    fn test_escher_record_creation() {
        let data_vec = vec![1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16];
        let record = EscherRecord {
            record_type: EscherRecordType::ClientAnchor,
            version: 1,
            instance: 0,
            data_length: 16,
            data: Cow::Owned(data_vec),
            children: Vec::new(),
            properties: Vec::new(),
        };

        assert_eq!(record.record_type, EscherRecordType::ClientAnchor);
        assert_eq!(record.version, 1);
        assert_eq!(record.data_length, 16);
        assert_eq!(record.data.len(), 16);
        assert!(record.properties.is_empty());
    }

    #[test]
    fn test_escher_record_type_conversion() {
        assert_eq!(
            EscherRecordType::from(0xF000),
            EscherRecordType::DggContainer
        );
        assert_eq!(
            EscherRecordType::from(0xF004),
            EscherRecordType::SpContainer
        );
        assert_eq!(
            EscherRecordType::from(0xF010),
            EscherRecordType::ClientAnchor
        );
        assert_eq!(
            EscherRecordType::from(0xF00D),
            EscherRecordType::ClientTextbox
        );
        assert_eq!(EscherRecordType::from(999), EscherRecordType::Unknown);
    }

    #[test]
    fn extracts_shadow_properties_with_spec_ids_and_full_width_values() {
        let record = EscherRecord {
            record_type: EscherRecordType::Opt,
            version: 3,
            instance: 6,
            data_length: 0,
            data: Cow::Borrowed(&[]),
            children: Vec::new(),
            properties: vec![
                EscherProperty::new(0x0200, 0),
                EscherProperty::new(0x0201, 0x0080_8080),
                EscherProperty::new(0x0204, 0x0000_8000),
                EscherProperty::new(0x0205, (-25_400_i32) as u32),
                EscherProperty::new(0x0206, 25_400),
                EscherProperty::new(0x023F, 0x0002_0002),
            ],
        };

        let values = record.extract_property_values();
        assert_eq!(values.shadow_type, Some(0));
        assert_eq!(values.shadow_color, Some(0x0080_8080));
        assert_eq!(values.shadow_opacity, Some(0x0000_8000));
        assert_eq!(values.shadow_offset_x, Some(-25_400));
        assert_eq!(values.shadow_offset_y, Some(25_400));
        assert_eq!(values.shadow_enabled, Some(true));
        assert_eq!(values.shadow_obscured, None);
    }

    #[test]
    fn extracts_common_properties_with_spec_ids_and_full_width_values() {
        let record = EscherRecord {
            record_type: EscherRecordType::Opt,
            version: 3,
            instance: 17,
            data_length: 0,
            data: Cow::Borrowed(&[]),
            children: Vec::new(),
            properties: vec![
                EscherProperty::new(0x0004, (-90_i32 * 65_536) as u32),
                EscherProperty::new(0x007F, 0x0080_0080),
                EscherProperty::new(0x0081, (-12_700_i32) as u32),
                EscherProperty::new(0x0082, 25_400),
                EscherProperty::new(0x0083, 38_100),
                EscherProperty::new(0x0084, 50_800),
                EscherProperty::new(0x0087, 4),
                EscherProperty::new(0x0180, 4),
                EscherProperty::new(0x0181, 0x0012_3456),
                EscherProperty::new(0x0182, 65_536),
                EscherProperty::new(0x0183, 0x0065_4321),
                EscherProperty::new(0x01C0, 0x00AB_CDEF),
                EscherProperty::new(0x01C1, 32_768),
                EscherProperty::new(0x01CB, 25_400),
                EscherProperty::new(0x01CD, 1),
                EscherProperty::new(0x01CE, 2),
            ],
        };

        let values = record.extract_property_values();
        assert_eq!(values.rotation, Some(-90 * 65_536));
        assert_eq!(values.lock_aspect_ratio, Some(true));
        assert_eq!(values.text_left_margin, Some(-12_700));
        assert_eq!(values.text_top_margin, Some(25_400));
        assert_eq!(values.text_right_margin, Some(38_100));
        assert_eq!(values.text_bottom_margin, Some(50_800));
        assert_eq!(values.text_anchor, Some(4));
        assert_eq!(values.fill_type, Some(4));
        assert_eq!(values.fill_color, Some(0x0012_3456));
        assert_eq!(values.fill_opacity, Some(65_536));
        assert_eq!(values.fill_back_color, Some(0x0065_4321));
        assert_eq!(values.line_color, Some(0x00AB_CDEF));
        assert_eq!(values.line_opacity, Some(32_768));
        assert_eq!(values.line_width, Some(25_400));
        assert_eq!(values.line_style, Some(1));
        assert_eq!(values.line_dash_style, Some(2));

        let shape = record.extract_shape_properties().unwrap();
        assert_eq!(shape.fill_color, Some(0x0012_3456));
        assert_eq!(shape.line_color, Some(0x00AB_CDEF));
        assert_eq!(shape.line_width, Some(2));
    }

    #[test]
    fn test_shape_properties_extraction() {
        let anchor_data = vec![
            0x20, 0x00, // top = 32 master units
            0x10, 0x00, // left = 16 master units
            0x74, 0x00, // right = 116 master units
            0x52, 0x00, // bottom = 82 master units
        ];
        let anchor_record = EscherRecord {
            record_type: EscherRecordType::ClientAnchor,
            version: 1,
            instance: 0,
            data_length: 8,
            data: Cow::Owned(anchor_data),
            children: Vec::new(),
            properties: Vec::new(),
        };

        let shape_props_data = vec![1, 0, 0, 0, 0, 0, 0, 0];
        let shape_props_record = EscherRecord {
            record_type: EscherRecordType::Sp,
            version: 1,
            instance: 1,
            data_length: 8,
            data: Cow::Owned(shape_props_data),
            children: Vec::new(),
            properties: Vec::new(),
        };

        // Create container record
        let container = EscherRecord {
            record_type: EscherRecordType::SpContainer,
            version: 0xF,
            instance: 0,
            data_length: 0,
            data: Cow::Owned(Vec::new()),
            children: vec![anchor_record, shape_props_record],
            properties: Vec::new(),
        };

        let props = container.extract_shape_properties().unwrap();
        assert_eq!(props.x, ppt_master_i64_to_emu_i32(16));
        assert_eq!(props.y, ppt_master_i64_to_emu_i32(32));
        assert_eq!(props.width, ppt_master_i64_to_emu_i32(100));
        assert_eq!(props.height, ppt_master_i64_to_emu_i32(50));
        assert_eq!(props.shape_type, ShapeType::AutoShape);
        assert_eq!(props.id, 1);
    }

    #[test]
    fn test_text_extraction() {
        let mut text_data = Vec::new();
        text_data.extend_from_slice(&0u16.to_le_bytes());
        text_data.extend_from_slice(&4000u16.to_le_bytes());
        text_data.extend_from_slice(&10u32.to_le_bytes());
        text_data.extend_from_slice(&[0x48, 0x00, 0x65, 0x00, 0x6C, 0x00, 0x6C, 0x00, 0x6F, 0x00]);
        let text_record = EscherRecord {
            record_type: EscherRecordType::ClientTextbox,
            version: 1,
            instance: 0,
            data_length: text_data.len() as u32,
            data: Cow::Owned(text_data),
            children: Vec::new(),
            properties: Vec::new(),
        };

        let container = EscherRecord {
            record_type: EscherRecordType::SpContainer,
            version: 0xF,
            instance: 0,
            data_length: 0,
            data: Cow::Owned(Vec::new()),
            children: vec![text_record],
            properties: Vec::new(),
        };

        let text = container.extract_text().unwrap();
        // Text record contains "Hello" followed by null terminator
        assert_eq!(text, "Hello");
    }
}
