use super::consts::*;
use super::file::{OleError, OleFile};
use std::collections::HashMap;
use std::io::{Read, Seek};

/// Metadata extracted from OLE property streams
///
/// This struct contains standard properties from SummaryInformation
/// and DocumentSummaryInformation streams.
#[derive(Debug, Default)]
pub struct OleMetadata {
    // SummaryInformation properties
    pub codepage: Option<u32>,
    pub title: Option<String>,
    pub subject: Option<String>,
    pub author: Option<String>,
    pub keywords: Option<String>,
    pub comments: Option<String>,
    pub template: Option<String>,
    pub last_saved_by: Option<String>,
    pub revision_number: Option<String>,
    pub create_time: Option<u64>,
    pub last_saved_time: Option<u64>,
    pub num_pages: Option<u32>,
    pub num_words: Option<u32>,
    pub num_chars: Option<u32>,
    pub creating_application: Option<String>,
    pub security: Option<u32>,

    // DocumentSummaryInformation properties
    pub category: Option<String>,
    pub manager: Option<String>,
    pub company: Option<String>,
}

/// Property value types
#[derive(Debug, Clone)]
pub enum PropertyValue {
    I2(i16),
    I4(i32),
    UI2(u16),
    UI4(u32),
    Bool(bool),
    Lpstr(String),
    Lpwstr(String),
    Filetime(u64),
    Blob(Vec<u8>),
    Empty,
}

impl<R: Read + Seek> OleFile<R> {
    /// Parse metadata from standard property streams
    ///
    /// This method attempts to parse SummaryInformation and
    /// DocumentSummaryInformation streams to extract metadata.
    pub fn get_metadata(&mut self) -> Result<OleMetadata, OleError> {
        let mut metadata = OleMetadata::default();

        // Try to parse SummaryInformation stream
        if let Ok(data) = self.open_stream(&["\u{0005}SummaryInformation"])
            && let Ok(props) = parse_property_stream(&data) {
            extract_summary_info(&mut metadata, &props);
        }

        // Try to parse DocumentSummaryInformation stream
        if let Ok(data) = self.open_stream(&["\u{0005}DocumentSummaryInformation"])
            && let Ok(props) = parse_property_stream(&data) {
            extract_document_summary_info(&mut metadata, &props);
        }

        Ok(metadata)
    }
}

/// Parse a property stream and return properties as a HashMap
///
/// Property streams contain metadata in a structured format according
/// to [MS-OLEPS] specification.
fn parse_property_stream(data: &[u8]) -> Result<HashMap<u32, PropertyValue>, OleError> {
    if data.len() < 48 {
        return Err(OleError::InvalidFormat(
            "Property stream too short".to_string(),
        ));
    }

    let mut properties = HashMap::new();

    // Skip header (28 bytes) and format ID (20 bytes)
    let section_offset = u32::from_le_bytes([data[44], data[45], data[46], data[47]]) as usize;

    if section_offset + 8 > data.len() {
        return Err(OleError::InvalidFormat(
            "Invalid section offset".to_string(),
        ));
    }

    // Read property count (section size at offset 0 is not used)
    let num_props = u32::from_le_bytes([
        data[section_offset + 4],
        data[section_offset + 5],
        data[section_offset + 6],
        data[section_offset + 7],
    ]);

    // Limit properties to prevent DoS
    let num_props = num_props.min(1000);

    // Parse each property
    for i in 0..num_props {
        let prop_offset = section_offset + 8 + (i as usize) * 8;
        if prop_offset + 8 > data.len() {
            break;
        }

        // Property ID
        let prop_id = u32::from_le_bytes([
            data[prop_offset],
            data[prop_offset + 1],
            data[prop_offset + 2],
            data[prop_offset + 3],
        ]);

        // Offset to property value
        let value_offset = section_offset
            + u32::from_le_bytes([
                data[prop_offset + 4],
                data[prop_offset + 5],
                data[prop_offset + 6],
                data[prop_offset + 7],
            ]) as usize;

        if value_offset + 4 > data.len() {
            continue;
        }

        // Property type
        let prop_type = u16::from_le_bytes([data[value_offset], data[value_offset + 1]]);

        // Parse property value based on type
        if let Ok(value) = parse_property_value(data, value_offset + 4, prop_type) {
            properties.insert(prop_id, value);
        }
    }

    Ok(properties)
}

/// Parse a single property value based on its type
fn parse_property_value(
    data: &[u8],
    offset: usize,
    prop_type: u16,
) -> Result<PropertyValue, OleError> {
    match prop_type {
        VT_I2 => {
            // 16-bit signed integer
            if offset + 2 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let value = i16::from_le_bytes([data[offset], data[offset + 1]]);
            Ok(PropertyValue::I2(value))
        }
        VT_I4 | VT_INT | VT_ERROR => {
            // 32-bit signed integer
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let value = i32::from_le_bytes([
                data[offset],
                data[offset + 1],
                data[offset + 2],
                data[offset + 3],
            ]);
            Ok(PropertyValue::I4(value))
        }
        VT_UI2 => {
            // 16-bit unsigned integer
            if offset + 2 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let value = u16::from_le_bytes([data[offset], data[offset + 1]]);
            Ok(PropertyValue::UI2(value))
        }
        VT_UI4 | VT_UINT => {
            // 32-bit unsigned integer
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let value = u32::from_le_bytes([
                data[offset],
                data[offset + 1],
                data[offset + 2],
                data[offset + 3],
            ]);
            Ok(PropertyValue::UI4(value))
        }
        VT_LPSTR | VT_BSTR => {
            // Code page string
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let str_len = u32::from_le_bytes([
                data[offset],
                data[offset + 1],
                data[offset + 2],
                data[offset + 3],
            ]) as usize;

            if offset + 4 + str_len > data.len() {
                return Err(OleError::InvalidFormat("String overflow".to_string()));
            }

            let str_bytes = &data[offset + 4..offset + 4 + str_len];
            // Remove null terminators
            let s = String::from_utf8_lossy(str_bytes)
                .trim_end_matches('\0')
                .to_string();
            Ok(PropertyValue::Lpstr(s))
        }
        VT_LPWSTR => {
            // Unicode string (UTF-16LE)
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let char_count = u32::from_le_bytes([
                data[offset],
                data[offset + 1],
                data[offset + 2],
                data[offset + 3],
            ]) as usize;

            let byte_len = char_count * 2;
            if offset + 4 + byte_len > data.len() {
                return Err(OleError::InvalidFormat("String overflow".to_string()));
            }

            // Decode UTF-16LE
            let mut utf16_chars = Vec::new();
            for i in 0..char_count {
                let byte_offset = offset + 4 + i * 2;
                let code_unit = u16::from_le_bytes([data[byte_offset], data[byte_offset + 1]]);
                if code_unit == 0 {
                    break;
                }
                utf16_chars.push(code_unit);
            }

            let s = String::from_utf16_lossy(&utf16_chars);
            Ok(PropertyValue::Lpwstr(s))
        }
        VT_FILETIME => {
            // 64-bit file time
            if offset + 8 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let low = u32::from_le_bytes([
                data[offset],
                data[offset + 1],
                data[offset + 2],
                data[offset + 3],
            ]) as u64;
            let high = u32::from_le_bytes([
                data[offset + 4],
                data[offset + 5],
                data[offset + 6],
                data[offset + 7],
            ]) as u64;
            let filetime = low | (high << 32);
            Ok(PropertyValue::Filetime(filetime))
        }
        VT_BOOL => {
            // Boolean (16-bit)
            if offset + 2 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let value = u16::from_le_bytes([data[offset], data[offset + 1]]);
            Ok(PropertyValue::Bool(value != 0))
        }
        VT_BLOB => {
            // Binary data
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let blob_len = u32::from_le_bytes([
                data[offset],
                data[offset + 1],
                data[offset + 2],
                data[offset + 3],
            ]) as usize;

            if offset + 4 + blob_len > data.len() {
                return Err(OleError::InvalidFormat("Blob overflow".to_string()));
            }

            let blob = data[offset + 4..offset + 4 + blob_len].to_vec();
            Ok(PropertyValue::Blob(blob))
        }
        VT_EMPTY | VT_NULL => Ok(PropertyValue::Empty),
        _ => {
            // Unsupported type
            Ok(PropertyValue::Empty)
        }
    }
}

/// Extract SummaryInformation properties into metadata
fn extract_summary_info(metadata: &mut OleMetadata, props: &HashMap<u32, PropertyValue>) {
    // Property IDs for SummaryInformation (start at 1)
    // 1: CODEPAGE
    if let Some(PropertyValue::UI2(v)) = props.get(&1) {
        metadata.codepage = Some(*v as u32);
    }

    // 2: TITLE
    if let Some(v) = props.get(&2) {
        metadata.title = extract_string(v);
    }

    // 3: SUBJECT
    if let Some(v) = props.get(&3) {
        metadata.subject = extract_string(v);
    }

    // 4: AUTHOR
    if let Some(v) = props.get(&4) {
        metadata.author = extract_string(v);
    }

    // 5: KEYWORDS
    if let Some(v) = props.get(&5) {
        metadata.keywords = extract_string(v);
    }

    // 6: COMMENTS
    if let Some(v) = props.get(&6) {
        metadata.comments = extract_string(v);
    }

    // 7: TEMPLATE
    if let Some(v) = props.get(&7) {
        metadata.template = extract_string(v);
    }

    // 8: LAST_SAVED_BY
    if let Some(v) = props.get(&8) {
        metadata.last_saved_by = extract_string(v);
    }

    // 9: REVISION_NUMBER
    if let Some(v) = props.get(&9) {
        metadata.revision_number = extract_string(v);
    }

    // 12: CREATE_TIME
    if let Some(PropertyValue::Filetime(v)) = props.get(&12) {
        metadata.create_time = Some(*v);
    }

    // 13: LAST_SAVED_TIME
    if let Some(PropertyValue::Filetime(v)) = props.get(&13) {
        metadata.last_saved_time = Some(*v);
    }

    // 14: NUM_PAGES
    if let Some(PropertyValue::I4(v)) = props.get(&14) {
        metadata.num_pages = Some(*v as u32);
    }

    // 15: NUM_WORDS
    if let Some(PropertyValue::I4(v)) = props.get(&15) {
        metadata.num_words = Some(*v as u32);
    }

    // 16: NUM_CHARS
    if let Some(PropertyValue::I4(v)) = props.get(&16) {
        metadata.num_chars = Some(*v as u32);
    }

    // 18: CREATING_APPLICATION
    if let Some(v) = props.get(&18) {
        metadata.creating_application = extract_string(v);
    }

    // 19: SECURITY
    if let Some(PropertyValue::I4(v)) = props.get(&19) {
        metadata.security = Some(*v as u32);
    }
}

/// Extract DocumentSummaryInformation properties into metadata
fn extract_document_summary_info(metadata: &mut OleMetadata, props: &HashMap<u32, PropertyValue>) {
    // 2: CATEGORY
    if let Some(v) = props.get(&2) {
        metadata.category = extract_string(v);
    }

    // 14: MANAGER
    if let Some(v) = props.get(&14) {
        metadata.manager = extract_string(v);
    }

    // 15: COMPANY
    if let Some(v) = props.get(&15) {
        metadata.company = extract_string(v);
    }
}

/// Extract string from property value
fn extract_string(value: &PropertyValue) -> Option<String> {
    match value {
        PropertyValue::Lpstr(s) | PropertyValue::Lpwstr(s) => {
            if s.is_empty() {
                None
            } else {
                Some(s.clone())
            }
        }
        _ => None,
    }
}
