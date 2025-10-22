use super::consts::*;
use super::file::{OleError, OleFile};
use chrono::{DateTime, Duration, Utc};
use std::collections::HashMap;
use std::io::{Read, Seek};
use zerocopy::{FromBytes, I16, I32, LE, U16, U32};

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
    pub edit_time: Option<Duration>,
    pub create_time: Option<DateTime<Utc>>,
    pub last_printed_time: Option<DateTime<Utc>>,
    pub last_saved_time: Option<DateTime<Utc>>,
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
    Lpstr(Vec<u8>), // Raw bytes for ANSI strings (need codepage to decode)
    Lpwstr(String), // Already decoded UTF-16
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
            && let Ok(props) = parse_property_stream(&data)
        {
            extract_summary_info(&mut metadata, &props);
        }

        // Try to parse DocumentSummaryInformation stream
        if let Ok(data) = self.open_stream(&["\u{0005}DocumentSummaryInformation"])
            && let Ok(props) = parse_property_stream(&data)
        {
            extract_document_summary_info(&mut metadata, &props);
        }

        Ok(metadata)
    }
}

/// Convert a FILETIME property value to Rust Date
///
/// The FILETIME structure is a 64-bit value that represents the number of 100-nanosecond intervals
/// that have elapsed since January 1, 1601, Coordinated Universal Time (UTC).
#[inline]
fn filetime_to_date(filetime: u64) -> Option<DateTime<Utc>> {
    // Number of 100-nanosecond intervals between 1601-01-01 and 1970-01-01
    const EPOCH_DIFF: i64 = 116_444_736_000_000_000;
    let doc_epoch = i64::try_from(filetime).ok()?;
    Some(DateTime::from_timestamp_nanos(
        (doc_epoch - EPOCH_DIFF) * 100,
    ))
}

/// Convert a FILETIME property value to Rust duration
///
/// It is like [filetime_to_date], but the result is a duration instead of a date.
#[inline]
fn filetime_to_duration(filetime: u64) -> Option<Duration> {
    let nanos = filetime * 100;
    Some(Duration::nanoseconds(i64::try_from(nanos).ok()?))
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

    // Skip header (28 bytes) and format ID (20 bytes)
    let section_offset = U32::<LE>::read_from_bytes(&data[44..48])
        .map(|v| v.get() as usize)
        .unwrap_or(0);

    if section_offset + 8 > data.len() {
        return Err(OleError::InvalidFormat(
            "Invalid section offset".to_string(),
        ));
    }

    // Read property count (section size at offset 0 is not used)
    let num_props = U32::<LE>::read_from_bytes(&data[section_offset + 4..section_offset + 8])
        .map(|v| v.get())
        .unwrap_or(0);

    // Limit properties to prevent DoS
    let num_props = num_props.min(1000);

    // Create a HashMap with the estimated number of properties
    let mut properties = HashMap::with_capacity(num_props as usize);

    // Parse each property
    for i in 0..num_props {
        let prop_offset = section_offset + 8 + (i as usize) * 8;
        if prop_offset + 8 > data.len() {
            break;
        }

        // Property ID
        let prop_id = U32::<LE>::read_from_bytes(&data[prop_offset..prop_offset + 4])
            .map(|v| v.get())
            .unwrap_or(0);

        // Offset to property value
        let value_offset = section_offset
            + U32::<LE>::read_from_bytes(&data[prop_offset + 4..prop_offset + 8])
                .map(|v| v.get() as usize)
                .unwrap_or(0);

        if value_offset + 4 > data.len() {
            continue;
        }

        // Property type
        let prop_type = U16::<LE>::read_from_bytes(&data[value_offset..value_offset + 2])
            .map(|v| v.get())
            .unwrap_or(0);

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
            let value = I16::<LE>::read_from_bytes(&data[offset..offset + 2])
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::I2(value))
        },
        VT_I4 | VT_INT | VT_ERROR => {
            // 32-bit signed integer
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let value = I32::<LE>::read_from_bytes(&data[offset..offset + 4])
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::I4(value))
        },
        VT_UI2 => {
            // 16-bit unsigned integer
            if offset + 2 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let value = U16::<LE>::read_from_bytes(&data[offset..offset + 2])
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::UI2(value))
        },
        VT_UI4 | VT_UINT => {
            // 32-bit unsigned integer
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let value = U32::<LE>::read_from_bytes(&data[offset..offset + 4])
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::UI4(value))
        },
        VT_LPSTR | VT_BSTR => {
            // Code page string
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let str_len = U32::<LE>::read_from_bytes(&data[offset..offset + 4])
                .map(|v| v.get() as usize)
                .unwrap_or(0);

            if offset + 4 + str_len > data.len() {
                return Err(OleError::InvalidFormat("String overflow".to_string()));
            }

            let str_bytes = &data[offset + 4..offset + 4 + str_len];
            // Store raw bytes - will be decoded later with proper codepage
            let raw_bytes = str_bytes.to_vec();
            Ok(PropertyValue::Lpstr(raw_bytes))
        },
        VT_LPWSTR => {
            // Unicode string (UTF-16LE)
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let char_count = U32::<LE>::read_from_bytes(&data[offset..offset + 4])
                .map(|v| v.get() as usize)
                .unwrap_or(0);

            let byte_len = char_count * 2;
            if offset + 4 + byte_len > data.len() {
                return Err(OleError::InvalidFormat("String overflow".to_string()));
            }

            // Decode UTF-16LE
            let mut utf16_chars = Vec::new();
            for i in 0..char_count {
                let byte_offset = offset + 4 + i * 2;
                let code_unit = U16::<LE>::read_from_bytes(&data[byte_offset..byte_offset + 2])
                    .map(|v| v.get())
                    .unwrap_or(0);
                if code_unit == 0 {
                    break;
                }
                utf16_chars.push(code_unit);
            }

            let s = String::from_utf16_lossy(&utf16_chars);
            Ok(PropertyValue::Lpwstr(s))
        },
        VT_FILETIME => {
            // 64-bit file time
            if offset + 8 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let low = U32::<LE>::read_from_bytes(&data[offset..offset + 4])
                .map(|v| v.get() as u64)
                .unwrap_or(0);
            let high = U32::<LE>::read_from_bytes(&data[offset + 4..offset + 8])
                .map(|v| v.get() as u64)
                .unwrap_or(0);
            let filetime = low | (high << 32);
            Ok(PropertyValue::Filetime(filetime))
        },
        VT_BOOL => {
            // Boolean (16-bit)
            if offset + 2 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let value = U16::<LE>::read_from_bytes(&data[offset..offset + 2])
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::Bool(value != 0))
        },
        VT_BLOB => {
            // Binary data
            if offset + 4 > data.len() {
                return Err(OleError::InvalidFormat("Buffer overflow".to_string()));
            }
            let blob_len = U32::<LE>::read_from_bytes(&data[offset..offset + 4])
                .map(|v| v.get() as usize)
                .unwrap_or(0);

            if offset + 4 + blob_len > data.len() {
                return Err(OleError::InvalidFormat("Blob overflow".to_string()));
            }

            let blob = data[offset + 4..offset + 4 + blob_len].to_vec();
            Ok(PropertyValue::Blob(blob))
        },
        VT_EMPTY | VT_NULL => Ok(PropertyValue::Empty),
        _ => {
            // Unsupported type
            Ok(PropertyValue::Empty)
        },
    }
}

/// Extract SummaryInformation properties into metadata
///
/// See [this document](https://learn.microsoft.com/en-us/openspecs/windows_protocols/MS-OLEPS/f7933d28-2cc4-4b36-bc23-8861cbcd37c4)
/// for your information.
fn extract_summary_info(metadata: &mut OleMetadata, props: &HashMap<u32, PropertyValue>) {
    // Property IDs for SummaryInformation (start at 1)
    // 1: CODEPAGE
    let codepage = if let Some(PropertyValue::I2(v)) = props.get(&1) {
        let cp = Some(*v as u32);
        metadata.codepage = cp;
        cp
    } else {
        None
    };

    // 2: TITLE
    if let Some(v) = props.get(&2) {
        metadata.title = extract_string(v, codepage);
    }

    // 3: SUBJECT
    if let Some(v) = props.get(&3) {
        metadata.subject = extract_string(v, codepage);
    }

    // 4: AUTHOR
    if let Some(v) = props.get(&4) {
        metadata.author = extract_string(v, codepage);
    }

    // 5: KEYWORDS
    if let Some(v) = props.get(&5) {
        metadata.keywords = extract_string(v, codepage);
    }

    // 6: COMMENTS
    if let Some(v) = props.get(&6) {
        metadata.comments = extract_string(v, codepage);
    }

    // 7: TEMPLATE
    if let Some(v) = props.get(&7) {
        metadata.template = extract_string(v, codepage);
    }

    // 8: LAST_SAVED_BY
    if let Some(v) = props.get(&8) {
        metadata.last_saved_by = extract_string(v, codepage);
    }

    // 9: REVISION_NUMBER
    if let Some(v) = props.get(&9) {
        metadata.revision_number = extract_string(v, codepage);
    }

    // 10: EDIT_TIME
    if let Some(PropertyValue::Filetime(v)) = props.get(&10) {
        metadata.edit_time = filetime_to_duration(*v);
    }

    // 11: LAST_PRINTED_TIME
    if let Some(PropertyValue::Filetime(v)) = props.get(&11) {
        metadata.last_printed_time = filetime_to_date(*v);
    }

    // 12: CREATE_TIME
    if let Some(PropertyValue::Filetime(v)) = props.get(&12) {
        metadata.create_time = filetime_to_date(*v);
    }

    // 13: LAST_SAVED_TIME
    if let Some(PropertyValue::Filetime(v)) = props.get(&13) {
        metadata.last_saved_time = filetime_to_date(*v);
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
        metadata.creating_application = extract_string(v, codepage);
    }

    // 19: SECURITY
    if let Some(PropertyValue::I4(v)) = props.get(&19) {
        metadata.security = Some(*v as u32);
    }
}

/// Extract DocumentSummaryInformation properties into metadata
///
/// See [this document](https://learn.microsoft.com/en-us/windows/win32/stg/the-documentsummaryinformation-and-userdefined-property-sets)
/// for your information.
fn extract_document_summary_info(metadata: &mut OleMetadata, props: &HashMap<u32, PropertyValue>) {
    // Use the codepage that was set during SummaryInformation parsing
    let codepage = metadata.codepage;

    // 2: CATEGORY
    if let Some(v) = props.get(&2) {
        metadata.category = extract_string(v, codepage);
    }

    // 3. PRESFORMAT
    // if let Some(v) = props.get(&3) {
    //     metadata.presentation_target = extract_string(v, codepage);
    // }

    // 14: MANAGER
    if let Some(v) = props.get(&14) {
        metadata.manager = extract_string(v, codepage);
    }

    // 15: COMPANY
    if let Some(v) = props.get(&15) {
        metadata.company = extract_string(v, codepage);
    }
}

/// Extract string from property value with proper encoding
fn extract_string(value: &PropertyValue, codepage: Option<u32>) -> Option<String> {
    match value {
        PropertyValue::Lpstr(bytes) => {
            if bytes.is_empty() {
                None
            } else {
                super::codepage::decode_bytes(bytes, codepage)
            }
        },
        PropertyValue::Lpwstr(s) => {
            if s.is_empty() {
                None
            } else {
                Some(s.clone())
            }
        },
        _ => None,
    }
}
