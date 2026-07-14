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
    let nanos = doc_epoch.checked_sub(EPOCH_DIFF)?.checked_mul(100)?;
    Some(DateTime::from_timestamp_nanos(nanos))
}

/// Convert a FILETIME property value to Rust duration
///
/// It is like [filetime_to_date], but the result is a duration instead of a date.
#[inline]
fn filetime_to_duration(filetime: u64) -> Option<Duration> {
    let nanos = filetime.checked_mul(100)?;
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
    if data[0..2] != [0xFE, 0xFF] {
        return Err(OleError::InvalidFormat(
            "Invalid property-set byte order".to_string(),
        ));
    }
    let property_set_count = read_u32(data, 24, "property-set count")?;
    if property_set_count == 0 || property_set_count > 2 {
        return Err(OleError::InvalidFormat(format!(
            "Invalid property-set count {property_set_count}"
        )));
    }
    let descriptors_len = usize::try_from(property_set_count)
        .ok()
        .and_then(|count| count.checked_mul(20))
        .and_then(|length| 28usize.checked_add(length))
        .ok_or_else(|| OleError::InvalidFormat("Property-set table overflow".to_string()))?;
    if descriptors_len > data.len() {
        return Err(OleError::InvalidFormat(
            "Truncated property-set table".to_string(),
        ));
    }

    // Skip header (28 bytes) and format ID (20 bytes)
    let section_offset = usize::try_from(read_u32(data, 44, "section offset")?)
        .map_err(|_| OleError::InvalidFormat("Section offset is too large".to_string()))?;

    checked_range(data, section_offset, 8, "section header")?;
    let section_size = usize::try_from(read_u32(data, section_offset, "section size")?)
        .map_err(|_| OleError::InvalidFormat("Section size is too large".to_string()))?;
    if section_size < 8 {
        return Err(OleError::InvalidFormat(
            "Property section is too short".to_string(),
        ));
    }
    let section_end = section_offset
        .checked_add(section_size)
        .filter(|end| *end <= data.len())
        .ok_or_else(|| OleError::InvalidFormat("Property section overflow".to_string()))?;

    let num_props = read_u32(data, section_offset + 4, "property count")?;

    if num_props > 1000 {
        return Err(OleError::InvalidFormat(
            "Property count exceeds the safety limit".to_string(),
        ));
    }
    let property_table_len = usize::try_from(num_props)
        .ok()
        .and_then(|count| count.checked_mul(8))
        .and_then(|length| 8usize.checked_add(length))
        .ok_or_else(|| OleError::InvalidFormat("Property table overflow".to_string()))?;
    if property_table_len > section_size {
        return Err(OleError::InvalidFormat(
            "Property table exceeds its section".to_string(),
        ));
    }

    // Create a HashMap with the estimated number of properties
    let mut properties = HashMap::with_capacity(num_props as usize);

    // Parse each property
    for i in 0..num_props {
        let prop_offset = section_offset + 8 + (i as usize) * 8;

        // Property ID
        let prop_id = read_u32(data, prop_offset, "property identifier")?;

        // Offset to property value
        let relative_offset =
            usize::try_from(read_u32(data, prop_offset + 4, "property value offset")?)
                .map_err(|_| OleError::InvalidFormat("Property offset is too large".to_string()))?;
        let value_offset = section_offset
            .checked_add(relative_offset)
            .ok_or_else(|| OleError::InvalidFormat("Property offset overflow".to_string()))?;

        if value_offset
            .checked_add(4)
            .is_none_or(|end| end > section_end)
        {
            return Err(OleError::InvalidFormat(
                "Property value exceeds its section".to_string(),
            ));
        }

        // Property type
        let prop_type = U16::<LE>::read_from_bytes(&data[value_offset..value_offset + 2])
            .map(|v| v.get())
            .unwrap_or(0);

        // Parse property value based on type
        if let Ok(value) = parse_property_value(&data[..section_end], value_offset + 4, prop_type) {
            properties.insert(prop_id, value);
        }
    }

    Ok(properties)
}

fn checked_range<'a>(
    data: &'a [u8],
    offset: usize,
    length: usize,
    description: &str,
) -> Result<&'a [u8], OleError> {
    let end = offset
        .checked_add(length)
        .filter(|end| *end <= data.len())
        .ok_or_else(|| OleError::InvalidFormat(format!("{description} overflow")))?;
    Ok(&data[offset..end])
}

fn read_u32(data: &[u8], offset: usize, description: &str) -> Result<u32, OleError> {
    let bytes = checked_range(data, offset, 4, description)?;
    U32::<LE>::read_from_bytes(bytes)
        .map(|value| value.get())
        .map_err(|_| OleError::InvalidFormat(format!("Invalid {description}")))
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
            let value = I16::<LE>::read_from_bytes(checked_range(data, offset, 2, "I2 value")?)
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::I2(value))
        },
        VT_I4 | VT_INT | VT_ERROR => {
            // 32-bit signed integer
            let value = I32::<LE>::read_from_bytes(checked_range(data, offset, 4, "I4 value")?)
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::I4(value))
        },
        VT_UI2 => {
            // 16-bit unsigned integer
            let value = U16::<LE>::read_from_bytes(checked_range(data, offset, 2, "UI2 value")?)
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::UI2(value))
        },
        VT_UI4 | VT_UINT => {
            // 32-bit unsigned integer
            let value = U32::<LE>::read_from_bytes(checked_range(data, offset, 4, "UI4 value")?)
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::UI4(value))
        },
        VT_LPSTR | VT_BSTR => {
            // Code page string
            let str_len = usize::try_from(read_u32(data, offset, "string length")?)
                .map_err(|_| OleError::InvalidFormat("String is too large".to_string()))?;
            let string_offset = offset
                .checked_add(4)
                .ok_or_else(|| OleError::InvalidFormat("String offset overflow".to_string()))?;
            let str_bytes = checked_range(data, string_offset, str_len, "string")?;
            // Store raw bytes - will be decoded later with proper codepage
            let raw_bytes = str_bytes.to_vec();
            Ok(PropertyValue::Lpstr(raw_bytes))
        },
        VT_LPWSTR => {
            // Unicode string (UTF-16LE)
            let char_count = usize::try_from(read_u32(data, offset, "Unicode string length")?)
                .map_err(|_| OleError::InvalidFormat("Unicode string is too large".to_string()))?;
            let byte_len = char_count.checked_mul(2).ok_or_else(|| {
                OleError::InvalidFormat("Unicode string length overflow".to_string())
            })?;
            let string_offset = offset.checked_add(4).ok_or_else(|| {
                OleError::InvalidFormat("Unicode string offset overflow".to_string())
            })?;
            let string_bytes = checked_range(data, string_offset, byte_len, "Unicode string")?;

            // Decode UTF-16LE
            let mut utf16_chars = Vec::with_capacity(char_count);
            for bytes in string_bytes.chunks_exact(2) {
                let code_unit = U16::<LE>::read_from_bytes(bytes)
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
            let bytes = checked_range(data, offset, 8, "FILETIME value")?;
            let low = U32::<LE>::read_from_bytes(&bytes[..4])
                .map(|v| v.get() as u64)
                .unwrap_or(0);
            let high = U32::<LE>::read_from_bytes(&bytes[4..])
                .map(|v| v.get() as u64)
                .unwrap_or(0);
            let filetime = low | (high << 32);
            Ok(PropertyValue::Filetime(filetime))
        },
        VT_BOOL => {
            // Boolean (16-bit)
            let value = U16::<LE>::read_from_bytes(checked_range(data, offset, 2, "BOOL value")?)
                .map(|v| v.get())
                .unwrap_or(0);
            Ok(PropertyValue::Bool(value != 0))
        },
        VT_BLOB => {
            // Binary data
            let blob_len = usize::try_from(read_u32(data, offset, "blob length")?)
                .map_err(|_| OleError::InvalidFormat("Blob is too large".to_string()))?;
            let blob_offset = offset
                .checked_add(4)
                .ok_or_else(|| OleError::InvalidFormat("Blob offset overflow".to_string()))?;
            let blob = checked_range(data, blob_offset, blob_len, "blob")?.to_vec();
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
        // PIDSI_CODEPAGE is stored as VT_I2, but code pages use all 16 bits
        // (for example UTF-8 is 65001, represented as a negative i16).
        let cp = Some(u32::from(*v as u16));
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
        metadata.num_pages = u32::try_from(*v).ok();
    }

    // 15: NUM_WORDS
    if let Some(PropertyValue::I4(v)) = props.get(&15) {
        metadata.num_words = u32::try_from(*v).ok();
    }

    // 16: NUM_CHARS
    if let Some(PropertyValue::I4(v)) = props.get(&16) {
        metadata.num_chars = u32::try_from(*v).ok();
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
    let codepage = match props.get(&1) {
        Some(PropertyValue::I2(value)) => Some(u32::from(*value as u16)),
        _ => metadata.codepage,
    };
    if metadata.codepage.is_none() {
        metadata.codepage = codepage;
    }

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
            let bytes = bytes.strip_suffix(&[0]).unwrap_or(bytes);
            if bytes.is_empty() {
                None
            } else {
                litchi_core::encoding::decode_bytes(bytes, codepage)
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

// Convert CFB-substrate metadata into the unified `litchi_core::Metadata`.
//
// Used to live in the umbrella crate's `metadata_ext.rs`, but the orphan rule
// forbids implementing `From<external> for external`. This impl is local to
// `litchi-cfb` because `OleMetadata` is defined here.
impl From<OleMetadata> for litchi_core::Metadata {
    fn from(ole_metadata: OleMetadata) -> Self {
        litchi_core::Metadata {
            title: ole_metadata.title,
            subject: ole_metadata.subject,
            author: ole_metadata.author,
            keywords: ole_metadata.keywords,
            description: ole_metadata.comments,
            template: ole_metadata.template,
            last_modified_by: ole_metadata.last_saved_by,
            revision: ole_metadata.revision_number,
            created: ole_metadata.create_time,
            modified: ole_metadata.last_saved_time,
            page_count: ole_metadata.num_pages,
            word_count: ole_metadata.num_words,
            character_count: ole_metadata.num_chars,
            application: ole_metadata.creating_application,
            category: ole_metadata.category,
            company: ole_metadata.company,
            manager: ole_metadata.manager,
            content_status: None, // OLE doesn't have this field
            last_printed_time: ole_metadata.last_printed_time,
            security: ole_metadata.security,
            codepage: ole_metadata.codepage,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn summary_property_stream() -> Vec<u8> {
        let mut data = vec![0u8; 96];
        data[0..2].copy_from_slice(&[0xFE, 0xFF]);
        data[24..28].copy_from_slice(&1u32.to_le_bytes());
        data[44..48].copy_from_slice(&48u32.to_le_bytes());

        data[48..52].copy_from_slice(&48u32.to_le_bytes());
        data[52..56].copy_from_slice(&2u32.to_le_bytes());
        data[56..60].copy_from_slice(&1u32.to_le_bytes());
        data[60..64].copy_from_slice(&24u32.to_le_bytes());
        data[64..68].copy_from_slice(&2u32.to_le_bytes());
        data[68..72].copy_from_slice(&32u32.to_le_bytes());

        data[72..74].copy_from_slice(&VT_I2.to_le_bytes());
        data[76..78].copy_from_slice(&65001u16.to_le_bytes());
        data[80..82].copy_from_slice(&VT_LPSTR.to_le_bytes());
        data[84..88].copy_from_slice(&6u32.to_le_bytes());
        data[88..94].copy_from_slice(b"Hello\0");
        data
    }

    #[test]
    fn parses_unsigned_codepage_and_trims_string_terminator() {
        let properties = parse_property_stream(&summary_property_stream()).unwrap();
        let mut metadata = OleMetadata::default();
        extract_summary_info(&mut metadata, &properties);
        assert_eq!(metadata.codepage, Some(65001));
        assert_eq!(metadata.title.as_deref(), Some("Hello"));
    }

    #[test]
    fn rejects_property_offsets_and_lengths_that_overflow() {
        let mut invalid_section = summary_property_stream();
        invalid_section[44..48].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(parse_property_stream(&invalid_section).is_err());

        let mut invalid_value = summary_property_stream();
        invalid_value[68..72].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(parse_property_stream(&invalid_value).is_err());

        let mut unicode = [0u8; 8];
        unicode[..4].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(parse_property_value(&unicode, 0, VT_LPWSTR).is_err());
    }

    #[test]
    fn rejects_invalid_property_set_headers_and_tables() {
        let mut invalid_byte_order = summary_property_stream();
        invalid_byte_order[0..2].copy_from_slice(&[0xFF, 0xFE]);
        assert!(parse_property_stream(&invalid_byte_order).is_err());

        let mut excessive_count = summary_property_stream();
        excessive_count[52..56].copy_from_slice(&1001u32.to_le_bytes());
        assert!(parse_property_stream(&excessive_count).is_err());

        let mut truncated_table = summary_property_stream();
        truncated_table[48..52].copy_from_slice(&8u32.to_le_bytes());
        assert!(parse_property_stream(&truncated_table).is_err());
    }

    #[test]
    fn filetime_conversion_is_checked() {
        const UNIX_EPOCH_FILETIME: u64 = 116_444_736_000_000_000;
        assert_eq!(
            filetime_to_date(UNIX_EPOCH_FILETIME).map(|date| date.timestamp()),
            Some(0)
        );
        assert_eq!(
            filetime_to_duration(10).map(|value| value.num_nanoseconds()),
            Some(Some(1000))
        );
        assert!(filetime_to_date(u64::MAX).is_none());
        assert!(filetime_to_duration(u64::MAX).is_none());
    }

    #[test]
    fn ignores_negative_document_statistics() {
        let properties = HashMap::from([
            (14, PropertyValue::I4(-1)),
            (15, PropertyValue::I4(-2)),
            (16, PropertyValue::I4(-3)),
        ]);
        let mut metadata = OleMetadata::default();
        extract_summary_info(&mut metadata, &properties);
        assert_eq!(metadata.num_pages, None);
        assert_eq!(metadata.num_words, None);
        assert_eq!(metadata.num_chars, None);
    }
}
