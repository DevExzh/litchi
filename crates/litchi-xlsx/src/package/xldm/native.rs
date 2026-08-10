//! Bounded, inert structural parsing for MS-XLDM section 2.3 native data files.
//!
//! Compressed segment and string buffers remain borrowed opaque bytes. This
//! module never decrypts, decompresses, evaluates, or materializes model data.

use std::error::Error;
use std::fmt;

use super::{GeneratedNameKind, Storage, classify_generated_path};

const MAX_NATIVE_FILES: usize = 65_536;
const MAX_NATIVE_ITEMS: usize = 1_048_576;
const MAX_STRING_PAGES: usize = 524_288;
const MAX_UNCOMPRESSED_STRING_PAGE_BYTES: u64 = 4_294_967_296;
const MAX_COMPRESSED_STRING_PAGE_BYTES: u64 = 536_870_912;
const STRING_PAGE_BEGIN_MARK: u32 = 0xAABB_CCDD;
const STRING_PAGE_END_MARK: u32 = 0xABCD_ABCD;
const HASH_MAGIC: u32 = 0x12B9_B6A5;

/// A structural MS-XLDM section 2.3 validation error.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct NativeError(String);

impl NativeError {
    fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}

impl fmt::Display for NativeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl Error for NativeError {}

/// Result type for section 2.3 inspection.
pub type NativeResult<T> = Result<T, NativeError>;

/// Whether the five common hash fields precede an `XM_TYPE_STRING` store.
///
/// Section 2.3.2.1.2 makes this dependent on `DictionaryFlags` in section
/// 2.5. Use `Auto` only when the byte layout has a unique interpretation.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum StringHashMode {
    Present,
    Absent,
    #[default]
    Auto,
}

/// Per-file information that section 2.3 delegates to later XML metadata.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct StringHashOverride {
    pub storage_path: String,
    pub mode: StringHashMode,
}

/// Options for storage-level section 2.3 inspection.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct NativeParseOptions {
    pub string_hash_overrides: Vec<StringHashOverride>,
}

impl NativeParseOptions {
    fn string_hash_mode(&self, path: &str) -> NativeResult<StringHashMode> {
        let mut result = StringHashMode::Auto;
        let mut found = false;
        for entry in &self.string_hash_overrides {
            if entry.storage_path == path {
                if found {
                    return Err(NativeError::new(format!(
                        "duplicate string dictionary override for {path}"
                    )));
                }
                found = true;
                result = entry.mode;
            }
        }
        Ok(result)
    }
}

/// A borrowed segment from an `.idf` file.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct IdfSegment<'a> {
    pub size_units: u64,
    pub bytes: &'a [u8],
}

/// The common section 2.3.1.1 layout shared by all `.idf` files.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct IdfFile<'a> {
    pub segments: Vec<IdfSegment<'a>>,
    pub trailing_zero_padding: &'a [u8],
}

/// The dictionary type stored in the first four bytes.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DictionaryType {
    Long,
    Real,
    String,
}

/// The common five hash fields from section 2.3.3.1.1.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct HashHeader {
    pub algorithm: i32,
    pub entry_size: u32,
    pub bin_size: u32,
    pub local_entry_count: u32,
    pub bin_count: i64,
}

/// A borrowed numeric dictionary vector.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct NumericDictionary<'a> {
    pub element_count: u64,
    pub element_size: u32,
    pub values: &'a [u8],
}

/// Huffman character-set mode for a compressed string page.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum HuffmanCharacterSetMode {
    Single,
    Multiple,
}

/// Borrowed page payload. Compressed buffers are deliberately not decoded.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum StringPageData<'a> {
    Uncompressed {
        remaining_store_characters: u64,
        used_characters: u64,
        allocation_size: u64,
        buffer: &'a [u8],
    },
    Compressed {
        total_bits: u32,
        character_set_mode: HuffmanCharacterSetMode,
        character_set: Option<u8>,
        allocation_size: u64,
        decode_bits: u32,
        encode_array: &'a [u8],
        buffer: &'a [u8],
    },
}

/// One page in an `XM_TYPE_STRING` dictionary.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct StringPage<'a> {
    pub contains_nulls: bool,
    pub start_index: u64,
    pub string_count: u64,
    pub data: StringPageData<'a>,
}

/// A compressed bit offset or uncompressed byte offset and its page ID.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct StringRecordHandle {
    pub offset: u32,
    pub page_id: u32,
}

/// A fully framed `XM_TYPE_STRING` dictionary.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct StringDictionary<'a> {
    pub string_count: u64,
    pub longest_string_characters: u64,
    pub pages: Vec<StringPage<'a>>,
    pub record_handles: Vec<StringRecordHandle>,
}

/// The typed body of a dictionary file.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum DictionaryBody<'a> {
    Numeric(NumericDictionary<'a>),
    String(StringDictionary<'a>),
}

/// A section 2.3.2 dictionary with borrowed value/string storage.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DictionaryFile<'a> {
    pub dictionary_type: DictionaryType,
    pub hash: Option<HashHeader>,
    pub body: DictionaryBody<'a>,
    pub trailing_zero_padding: &'a [u8],
}

/// Optional hash statistics from section 2.3.3.1.2.1.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct HashStatistics<'a> {
    pub element_count: u64,
    pub bin_count: u64,
    pub used_bin_count: u64,
    pub fast_access_elements: u64,
    pub locals_per_bin: u64,
    pub maximum_chain: u64,
    pub histogram_element_size: u32,
    pub histogram: &'a [u8],
}

/// A borrowed persisted hash bin and its used local entries.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct HashBin<'a> {
    pub entry_count: u32,
    pub local_entries: &'a [u8],
}

/// A section 2.3.3.1 hash index. Hash entries remain borrowed records.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct HashIndexFile<'a> {
    pub header: HashHeader,
    pub record_count: u64,
    pub current_mask: u64,
    pub statistics: Option<HashStatistics<'a>>,
    pub bins: Vec<HashBin<'a>>,
    pub collision_count: u64,
    pub collision_entries: &'a [u8],
    pub trailing_zero_padding: &'a [u8],
}

/// A recognized native data member.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum NativeData<'a> {
    Idf(IdfFile<'a>),
    Dictionary(DictionaryFile<'a>),
    HashIndex(HashIndexFile<'a>),
}

/// A generated storage member and its typed, borrowed section 2.3 view.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct NativeFile<'a> {
    pub storage_path: &'a str,
    pub bytes: &'a [u8],
    pub data: NativeData<'a>,
}

/// All section 2.3-compatible members found in a validated storage object.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct NativeModel<'a> {
    pub files: Vec<NativeFile<'a>>,
}

/// Inspect every generated `.idf`, `.dictionary`, and `.hidx` member logged by
/// a previously validated MS-XLDM storage object.
pub fn inspect<'a>(
    storage: &'a Storage<'a>,
    options: &NativeParseOptions,
) -> NativeResult<NativeModel<'a>> {
    let mut files = Vec::new();
    for group in &storage.backup_log.file_groups {
        for logged in &group.files {
            let path = logged.storage_path.as_str();
            let suffix = if path.ends_with(".dictionary") {
                Some("dictionary")
            } else if path.ends_with(".hidx") {
                Some("hidx")
            } else if path.ends_with(".idf") {
                Some("idf")
            } else {
                None
            };
            let Some(suffix) = suffix else { continue };
            if files.len() == MAX_NATIVE_FILES {
                return Err(NativeError::new("too many native data files"));
            }
            let generated = classify_generated_path(path).map_err(|error| {
                NativeError::new(format!("invalid generated path {path}: {error}"))
            })?;
            let directory_index = storage
                .files
                .iter()
                .position(|entry| entry.path == path)
                .ok_or_else(|| {
                    NativeError::new(format!(
                        "logged native member {path} is absent from the directory"
                    ))
                })?;
            let bytes = storage
                .file_payload(directory_index)
                .ok_or_else(|| NativeError::new(format!("cannot resolve native member {path}")))?;
            let data = match suffix {
                "dictionary" => NativeData::Dictionary(parse_dictionary(
                    bytes,
                    options.string_hash_mode(path)?,
                )?),
                "hidx" => NativeData::HashIndex(parse_hash_index(bytes)?),
                _ => NativeData::Idf(parse_storage_idf(generated.kind, bytes)?),
            };
            files.push(NativeFile {
                storage_path: path,
                bytes,
                data,
            });
        }
    }
    Ok(NativeModel { files })
}

fn parse_storage_idf(kind: GeneratedNameKind, bytes: &[u8]) -> NativeResult<IdfFile<'_>> {
    let idf = parse_idf(bytes)?;
    if kind == GeneratedNameKind::ColumnData
        && (idf.segments.len() < 2 || idf.segments.len() % 2 != 0)
    {
        return Err(NativeError::new(
            "column data requires primary/subsegment pairs",
        ));
    }
    Ok(idf)
}

/// Parse the common framing shared by all `.idf` files.
pub fn parse_idf(bytes: &[u8]) -> NativeResult<IdfFile<'_>> {
    let mut cursor = Cursor::new(bytes);
    let mut segments = Vec::new();
    while cursor.remaining() >= 8 {
        if segments.len() == MAX_NATIVE_ITEMS {
            return Err(NativeError::new("too many .idf segments"));
        }
        let size_units = cursor.read_u64("segment size")?;
        let byte_len = usize_from_u64(size_units, "segment size")?
            .checked_mul(8)
            .ok_or_else(|| NativeError::new("segment size overflow"))?;
        let segment = cursor.take(byte_len, "segment payload")?;
        segments.push(IdfSegment {
            size_units,
            bytes: segment,
        });
    }
    if segments.is_empty() {
        return Err(NativeError::new(
            "an .idf file requires at least one segment",
        ));
    }
    let padding = cursor.rest();
    require_zeroes(padding, ".idf trailing padding")?;
    Ok(IdfFile {
        segments,
        trailing_zero_padding: padding,
    })
}

/// Parse a dictionary file. `string_hash_mode` represents the section 2.5
/// `DictionaryFlags` choice when the dictionary type is `XM_TYPE_STRING`.
pub fn parse_dictionary(
    bytes: &[u8],
    string_hash_mode: StringHashMode,
) -> NativeResult<DictionaryFile<'_>> {
    let mut cursor = Cursor::new(bytes);
    let dictionary_type = match cursor.read_i32("dictionary type")? {
        0 => DictionaryType::Long,
        1 => DictionaryType::Real,
        2 => DictionaryType::String,
        value => {
            return Err(NativeError::new(format!("invalid dictionary type {value}")));
        },
    };
    match dictionary_type {
        DictionaryType::Long | DictionaryType::Real => {
            parse_numeric_dictionary(cursor, dictionary_type)
        },
        DictionaryType::String => match string_hash_mode {
            StringHashMode::Present => parse_string_dictionary(cursor, true),
            StringHashMode::Absent => parse_string_dictionary(cursor, false),
            StringHashMode::Auto => {
                let present = parse_string_dictionary(cursor.clone(), true);
                let absent = parse_string_dictionary(cursor, false);
                match (present, absent) {
                    (Ok(value), Err(_)) | (Err(_), Ok(value)) => Ok(value),
                    (Ok(_), Ok(_)) => Err(NativeError::new(
                        "ambiguous string dictionary hash layout; supply DictionaryFlags metadata",
                    )),
                    (Err(with_hash), Err(without_hash)) => Err(NativeError::new(format!(
                        "invalid string dictionary with hash ({with_hash}) and without hash ({without_hash})"
                    ))),
                }
            },
        },
    }
}

fn parse_numeric_dictionary(
    mut cursor: Cursor<'_>,
    dictionary_type: DictionaryType,
) -> NativeResult<DictionaryFile<'_>> {
    let hash = read_hash_header(&mut cursor)?;
    if hash.algorithm != -1 || hash.bin_count != -1 {
        return Err(NativeError::new(
            "numeric dictionaries require XM_INVALID algorithm and bin count",
        ));
    }
    validate_dictionary_hash_sizes(hash)?;
    let element_count = cursor.read_u64("dictionary element count")?;
    let element_size = cursor.read_u32("dictionary element size")?;
    let valid_size = match dictionary_type {
        DictionaryType::Long => element_size == 4 || element_size == 8,
        DictionaryType::Real => element_size == 8,
        DictionaryType::String => false,
    };
    if !valid_size {
        return Err(NativeError::new("invalid numeric dictionary element size"));
    }
    let values_len = bounded_product(element_count, u64::from(element_size), "dictionary vector")?;
    let values = cursor.take(values_len, "dictionary values")?;
    let padding = cursor.rest();
    require_zeroes(padding, "dictionary trailing padding")?;
    Ok(DictionaryFile {
        dictionary_type,
        hash: Some(hash),
        body: DictionaryBody::Numeric(NumericDictionary {
            element_count,
            element_size,
            values,
        }),
        trailing_zero_padding: padding,
    })
}

fn parse_string_dictionary(
    mut cursor: Cursor<'_>,
    has_hash: bool,
) -> NativeResult<DictionaryFile<'_>> {
    let hash = if has_hash {
        let hash = read_hash_header(&mut cursor)?;
        if !matches!(hash.algorithm, 0..=2) || hash.bin_count != -1 {
            return Err(NativeError::new(
                "string dictionary hash header has invalid algorithm or bin count",
            ));
        }
        validate_dictionary_hash_sizes(hash)?;
        Some(hash)
    } else {
        None
    };
    let string_count = cursor.read_u64("store string count")?;
    bound_count(string_count, MAX_NATIVE_ITEMS, "store string count")?;
    let store_compressed = cursor.read_bool("store compressed flag")?;
    let longest_string_characters = cursor.read_u64("longest string length")?;
    let page_count = cursor.read_u64("store page count")?;
    let page_count = bound_count(page_count, MAX_STRING_PAGES, "store page count")?;
    let mut pages = Vec::with_capacity(page_count);
    let mut cumulative_strings = 0u64;
    let mut any_compressed = false;
    for page_id in 0..page_count {
        let mask = cursor.read_u64("page mask")?;
        if mask > 1 {
            return Err(NativeError::new("invalid string page mask"));
        }
        let contains_nulls = cursor.read_bool("page NULL flag")?;
        let start_index = cursor.read_u64("page start index")?;
        if start_index != cumulative_strings {
            return Err(NativeError::new("noncontiguous page record-handle range"));
        }
        let page_string_count = cursor.read_u64("page string count")?;
        bound_count(page_string_count, MAX_NATIVE_ITEMS, "page string count")?;
        cumulative_strings = cumulative_strings
            .checked_add(page_string_count)
            .ok_or_else(|| NativeError::new("page string count overflow"))?;
        if cumulative_strings > string_count {
            return Err(NativeError::new("page strings exceed store string count"));
        }
        let compressed = cursor.read_bool("page compressed flag")?;
        if compressed != (mask == 1) {
            return Err(NativeError::new("page mask and compressed flag disagree"));
        }
        any_compressed |= compressed;
        if cursor.read_u32("string store begin mark")? != STRING_PAGE_BEGIN_MARK {
            return Err(NativeError::new("invalid string page begin mark"));
        }
        let data = if compressed {
            parse_compressed_page(&mut cursor)?
        } else {
            parse_uncompressed_page(&mut cursor)?
        };
        if cursor.read_u32("string store end mark")? != STRING_PAGE_END_MARK {
            return Err(NativeError::new("invalid string page end mark"));
        }
        let _ = page_id;
        pages.push(StringPage {
            contains_nulls,
            start_index,
            string_count: page_string_count,
            data,
        });
    }
    if cumulative_strings != string_count {
        return Err(NativeError::new(
            "page string counts do not cover the store",
        ));
    }
    if any_compressed != store_compressed {
        return Err(NativeError::new(
            "store compressed flag does not match its pages",
        ));
    }
    let handle_count = cursor.read_u64("record handle count")?;
    if handle_count != string_count {
        return Err(NativeError::new(
            "record handle count does not match string count",
        ));
    }
    if cursor.read_u32("record handle size")? != 8 {
        return Err(NativeError::new("record handle size must be 8"));
    }
    let handle_count_usize = bound_count(handle_count, MAX_NATIVE_ITEMS, "record handle count")?;
    let mut record_handles = Vec::with_capacity(handle_count_usize);
    for _ in 0..handle_count_usize {
        record_handles.push(StringRecordHandle {
            offset: cursor.read_u32("record handle offset")?,
            page_id: cursor.read_u32("record handle page ID")?,
        });
    }
    validate_record_handles(&pages, &record_handles)?;
    let padding = cursor.rest();
    require_zeroes(padding, "dictionary trailing padding")?;
    Ok(DictionaryFile {
        dictionary_type: DictionaryType::String,
        hash,
        body: DictionaryBody::String(StringDictionary {
            string_count,
            longest_string_characters,
            pages,
            record_handles,
        }),
        trailing_zero_padding: padding,
    })
}

fn parse_uncompressed_page<'a>(cursor: &mut Cursor<'a>) -> NativeResult<StringPageData<'a>> {
    let remaining_store_characters = cursor.read_u64("remaining store characters")?;
    let used_characters = cursor.read_u64("used store characters")?;
    let allocation_size = cursor.read_u64("uncompressed allocation size")?;
    if allocation_size > MAX_UNCOMPRESSED_STRING_PAGE_BYTES {
        return Err(NativeError::new(
            "uncompressed string page exceeds its byte limit",
        ));
    }
    let allocation = usize_from_u64(allocation_size, "uncompressed allocation size")?;
    if used_characters > allocation_size {
        return Err(NativeError::new(
            "used characters exceed the uncompressed allocation bound",
        ));
    }
    let buffer = cursor.take(allocation, "uncompressed string buffer")?;
    Ok(StringPageData::Uncompressed {
        remaining_store_characters,
        used_characters,
        allocation_size,
        buffer,
    })
}

fn parse_compressed_page<'a>(cursor: &mut Cursor<'a>) -> NativeResult<StringPageData<'a>> {
    let total_bits = cursor.read_u32("compressed store bit count")?;
    let character_set_mode = match cursor.read_u32("character set mode")? {
        703_121 => HuffmanCharacterSetMode::Single,
        703_122 => HuffmanCharacterSetMode::Multiple,
        _ => return Err(NativeError::new("invalid Huffman character set mode")),
    };
    let allocation_size = cursor.read_u64("compressed allocation size")?;
    if allocation_size > MAX_COMPRESSED_STRING_PAGE_BYTES {
        return Err(NativeError::new(
            "compressed string page exceeds its byte limit",
        ));
    }
    let character_set = match character_set_mode {
        HuffmanCharacterSetMode::Single => Some(cursor.read_u8("character set")?),
        HuffmanCharacterSetMode::Multiple => None,
    };
    let decode_bits = cursor.read_u32("Huffman decode bits")?;
    if !(2..=12).contains(&decode_bits) {
        return Err(NativeError::new("Huffman decode bits must be in 2..=12"));
    }
    let encode_array = cursor.take(128, "Huffman encode array")?;
    let buffer_size = cursor.read_u64("compressed buffer size")?;
    if buffer_size != allocation_size {
        return Err(NativeError::new(
            "compressed buffer and allocation sizes differ",
        ));
    }
    let buffer_len = usize_from_u64(buffer_size, "compressed buffer size")?;
    if u64::from(total_bits) > buffer_size.saturating_mul(8) {
        return Err(NativeError::new("compressed bit count exceeds its buffer"));
    }
    let buffer = cursor.take(buffer_len, "compressed string buffer")?;
    Ok(StringPageData::Compressed {
        total_bits,
        character_set_mode,
        character_set,
        allocation_size,
        decode_bits,
        encode_array,
        buffer,
    })
}

fn validate_record_handles(
    pages: &[StringPage<'_>],
    handles: &[StringRecordHandle],
) -> NativeResult<()> {
    for (page_id, page) in pages.iter().enumerate() {
        let start = usize_from_u64(page.start_index, "page start index")?;
        let count = usize_from_u64(page.string_count, "page string count")?;
        let end = start
            .checked_add(count)
            .ok_or_else(|| NativeError::new("page handle range overflow"))?;
        let page_handles = handles
            .get(start..end)
            .ok_or_else(|| NativeError::new("page handle range is out of bounds"))?;
        let mut previous = None;
        for handle in page_handles {
            if handle.page_id as usize != page_id {
                return Err(NativeError::new("record handle references the wrong page"));
            }
            if let Some(previous) = previous {
                if handle.offset <= previous {
                    return Err(NativeError::new(
                        "record handle offsets are not strictly increasing",
                    ));
                }
            } else if handle.offset != 0 {
                return Err(NativeError::new(
                    "the first record handle on a page must start at zero",
                ));
            }
            let limit = match &page.data {
                StringPageData::Uncompressed {
                    allocation_size, ..
                } => *allocation_size,
                StringPageData::Compressed { total_bits, .. } => u64::from(*total_bits),
            };
            if u64::from(handle.offset) >= limit && page.string_count != 0 {
                return Err(NativeError::new(
                    "record handle offset is out of page bounds",
                ));
            }
            previous = Some(handle.offset);
        }
    }
    Ok(())
}

/// Parse and verify a complete `.hidx` hash index without interpreting keys as
/// data values.
pub fn parse_hash_index(bytes: &[u8]) -> NativeResult<HashIndexFile<'_>> {
    let mut cursor = Cursor::new(bytes);
    let header = read_hash_header(&mut cursor)?;
    if header.algorithm != -1 {
        return Err(NativeError::new(
            "hash indexes require XM_INVALID algorithm",
        ));
    }
    if header.bin_count < 16 {
        return Err(NativeError::new("hash index bin count is below 16"));
    }
    let bin_count = usize_from_i64(header.bin_count, "hash bin count")?;
    if bin_count > MAX_NATIVE_ITEMS || !bin_count.is_power_of_two() {
        return Err(NativeError::new(
            "hash bin count is unbounded or not a power of two",
        ));
    }
    let entry_size = header.entry_size as usize;
    let bin_size = header.bin_size as usize;
    if !(8..=64).contains(&entry_size) || !matches!(bin_size, 64 | 128) {
        return Err(NativeError::new("invalid persisted hash entry or bin size"));
    }
    let expected_locals = (bin_size - 12) / entry_size;
    if header.local_entry_count as usize != expected_locals {
        return Err(NativeError::new(
            "local entry count does not match cache-aligned structure sizes",
        ));
    }
    let record_count = cursor.read_u64("hash record count")?;
    bound_count(record_count, MAX_NATIVE_ITEMS, "hash record count")?;
    let current_mask = cursor.read_u64("hash current mask")?;
    let wire_bin_count = u64::try_from(header.bin_count)
        .map_err(|_source| NativeError::new("hash bin count is negative"))?;
    if current_mask != wire_bin_count - 1 {
        return Err(NativeError::new(
            "hash current mask does not equal bins minus one",
        ));
    }
    let has_statistics = cursor.read_bool("hash statistics flag")?;
    let statistics = if has_statistics {
        Some(read_hash_statistics(&mut cursor)?)
    } else {
        None
    };
    let bin_count = usize::try_from(header.bin_count)
        .map_err(|_source| NativeError::new("hash bin count exceeds platform size"))?;
    let bins_len = bin_count
        .checked_mul(bin_size)
        .ok_or_else(|| NativeError::new("hash bins size overflow"))?;
    let bins_bytes = cursor.take(bins_len, "hash bins")?;
    let mut bins = Vec::with_capacity(bin_count);
    let mut summed_records = 0u64;
    let mut expected_collisions = 0u64;
    let mut used_bins = 0u64;
    let mut fast_access = 0u64;
    let mut maximum_chain = 0u64;
    let mut histogram_counts = vec![0u64; 1];
    for (index, raw_bin) in bins_bytes.chunks_exact(bin_size).enumerate() {
        require_zeroes(&raw_bin[..8], "persisted hash chain pointer")?;
        let count = u32::from_le_bytes(raw_bin[8..12].try_into().unwrap_or_else(|error| {
            crate::error::panic_error_invariant("operation was checked before extraction", error)
        }));
        let count64 = u64::from(count);
        summed_records = summed_records
            .checked_add(count64)
            .ok_or_else(|| NativeError::new("hash record total overflow"))?;
        maximum_chain = maximum_chain.max(count64);
        if count != 0 {
            used_bins += 1;
        }
        let local_count = (count as usize).min(expected_locals);
        fast_access += local_count as u64;
        expected_collisions += count64.saturating_sub(expected_locals as u64);
        if count as usize >= histogram_counts.len() {
            histogram_counts.resize(count as usize + 1, 0);
        }
        histogram_counts[count as usize] += 1;
        let local_len = local_count
            .checked_mul(entry_size)
            .ok_or_else(|| NativeError::new("local hash entry size overflow"))?;
        let local_entries = &raw_bin[12..12 + local_len];
        for entry in local_entries.chunks_exact(entry_size) {
            validate_hash_entry(entry, index, bin_count)?;
        }
        bins.push(HashBin {
            entry_count: count,
            local_entries,
        });
    }
    if summed_records != record_count {
        return Err(NativeError::new(
            "hash bin entry counts do not equal the record count",
        ));
    }
    let collision_count = cursor.read_u64("hash collision count")?;
    if collision_count != expected_collisions {
        return Err(NativeError::new(
            "hash collision count does not match bin overflows",
        ));
    }
    let collisions_len = bounded_product(
        collision_count,
        u64::from(header.entry_size),
        "collision entries",
    )?;
    let collision_entries = cursor.take(collisions_len, "collision entries")?;
    let mut collision_cursor = 0usize;
    for (bin_index, bin) in bins.iter().enumerate() {
        let overflow = (bin.entry_count as usize).saturating_sub(expected_locals);
        for _ in 0..overflow {
            let entry = &collision_entries[collision_cursor..collision_cursor + entry_size];
            validate_hash_entry(entry, bin_index, bin_count)?;
            collision_cursor += entry_size;
        }
    }
    if let Some(statistics) = &statistics {
        validate_hash_statistics(
            statistics,
            record_count,
            wire_bin_count,
            used_bins,
            fast_access,
            expected_locals as u64,
            maximum_chain,
            &histogram_counts,
        )?;
    }
    let padding = cursor.rest();
    require_zeroes(padding, "hash index trailing padding")?;
    Ok(HashIndexFile {
        header,
        record_count,
        current_mask,
        statistics,
        bins,
        collision_count,
        collision_entries,
        trailing_zero_padding: padding,
    })
}

fn read_hash_header(cursor: &mut Cursor<'_>) -> NativeResult<HashHeader> {
    Ok(HashHeader {
        algorithm: cursor.read_i32("hash algorithm")?,
        entry_size: cursor.read_u32("hash entry size")?,
        bin_size: cursor.read_u32("hash bin size")?,
        local_entry_count: cursor.read_u32("local entry count")?,
        bin_count: cursor.read_i64("hash bin count")?,
    })
}

fn validate_dictionary_hash_sizes(header: HashHeader) -> NativeResult<()> {
    if header.entry_size == 0
        || header.entry_size > 64
        || !matches!(header.bin_size, 64 | 128)
        || header.local_entry_count != (header.bin_size.saturating_sub(12) / header.entry_size)
    {
        return Err(NativeError::new(
            "dictionary hash structure sizes are inconsistent",
        ));
    }
    Ok(())
}

fn read_hash_statistics<'a>(cursor: &mut Cursor<'a>) -> NativeResult<HashStatistics<'a>> {
    let element_count = cursor.read_u64("statistics element count")?;
    let bin_count = cursor.read_u64("statistics bin count")?;
    let used_bin_count = cursor.read_u64("statistics used bin count")?;
    let fast_access_elements = cursor.read_u64("statistics fast access count")?;
    let locals_per_bin = cursor.read_u64("statistics locals per bin")?;
    let maximum_chain = cursor.read_u64("statistics maximum chain")?;
    let histogram_count = cursor.read_u64("histogram element count")?;
    bound_count(histogram_count, MAX_NATIVE_ITEMS, "histogram element count")?;
    let histogram_element_size = cursor.read_u32("histogram element size")?;
    if !matches!(histogram_element_size, 1 | 2 | 4 | 8) {
        return Err(NativeError::new("invalid histogram element size"));
    }
    let histogram_len = bounded_product(
        histogram_count,
        u64::from(histogram_element_size),
        "hash histogram",
    )?;
    let histogram = cursor.take(histogram_len, "hash histogram")?;
    Ok(HashStatistics {
        element_count,
        bin_count,
        used_bin_count,
        fast_access_elements,
        locals_per_bin,
        maximum_chain,
        histogram_element_size,
        histogram,
    })
}

#[allow(
    clippy::too_many_arguments,
    reason = "arguments mirror the fixed-width native XLDM record layout"
)]
fn validate_hash_statistics(
    statistics: &HashStatistics<'_>,
    records: u64,
    bins: u64,
    used_bins: u64,
    fast_access: u64,
    locals: u64,
    maximum_chain: u64,
    expected_histogram: &[u64],
) -> NativeResult<()> {
    if statistics.element_count != records
        || statistics.bin_count != bins
        || statistics.used_bin_count != used_bins
        || statistics.fast_access_elements != fast_access
        || statistics.locals_per_bin != locals
        || statistics.maximum_chain != maximum_chain
    {
        return Err(NativeError::new("hash statistics disagree with hash bins"));
    }
    let size = statistics.histogram_element_size as usize;
    let mut actual = Vec::new();
    for value in statistics.histogram.chunks_exact(size) {
        let mut bytes = [0u8; 8];
        bytes[..size].copy_from_slice(value);
        actual.push(u64::from_le_bytes(bytes));
    }
    if actual.len() < expected_histogram.len()
        || actual[..expected_histogram.len()] != *expected_histogram
        || actual[expected_histogram.len()..]
            .iter()
            .any(|value| *value != 0)
    {
        return Err(NativeError::new("hash histogram disagrees with hash bins"));
    }
    Ok(())
}

fn validate_hash_entry(entry: &[u8], bin_index: usize, bin_count: usize) -> NativeResult<()> {
    let persisted_hash = u32::from_le_bytes(entry[..4].try_into().unwrap_or_else(|error| {
        crate::error::panic_error_invariant("operation was checked before extraction", error)
    }));
    let key = u32::from_le_bytes(entry[4..8].try_into().unwrap_or_else(|error| {
        crate::error::panic_error_invariant("operation was checked before extraction", error)
    }));
    let bits = bin_count.trailing_zeros();
    let calculated = key
        .wrapping_mul(HASH_MAGIC)
        .checked_shr(32 - bits)
        .unwrap_or(0);
    if persisted_hash != calculated || calculated as usize != bin_index {
        return Err(NativeError::new(
            "hash entry value or bin placement is invalid",
        ));
    }
    Ok(())
}

/// Revalidate an inspected native member and return its exact original bytes.
pub fn write_file(file: &NativeFile<'_>) -> NativeResult<Vec<u8>> {
    match &file.data {
        NativeData::Idf(_) => {
            parse_idf(file.bytes)?;
        },
        NativeData::Dictionary(dictionary) => {
            let mode = if dictionary.hash.is_some() {
                StringHashMode::Present
            } else {
                StringHashMode::Absent
            };
            parse_dictionary(file.bytes, mode)?;
        },
        NativeData::HashIndex(_) => {
            parse_hash_index(file.bytes)?;
        },
    }
    Ok(file.bytes.to_vec())
}

#[derive(Clone)]
struct Cursor<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Cursor<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }

    fn remaining(&self) -> usize {
        self.bytes.len() - self.offset
    }

    fn take(&mut self, length: usize, field: &str) -> NativeResult<&'a [u8]> {
        let end = self
            .offset
            .checked_add(length)
            .ok_or_else(|| NativeError::new(format!("{field} range overflow")))?;
        let value = self
            .bytes
            .get(self.offset..end)
            .ok_or_else(|| NativeError::new(format!("truncated {field}")))?;
        self.offset = end;
        Ok(value)
    }

    fn read_u8(&mut self, field: &str) -> NativeResult<u8> {
        Ok(self.take(1, field)?[0])
    }

    fn read_bool(&mut self, field: &str) -> NativeResult<bool> {
        match self.read_u8(field)? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(NativeError::new(format!("invalid Boolean {field}"))),
        }
    }

    fn read_u32(&mut self, field: &str) -> NativeResult<u32> {
        Ok(u32::from_le_bytes(
            self.take(4, field)?.try_into().unwrap_or_else(|error| {
                crate::error::panic_error_invariant(
                    "operation was checked before extraction",
                    error,
                )
            }),
        ))
    }

    fn read_i32(&mut self, field: &str) -> NativeResult<i32> {
        Ok(i32::from_le_bytes(
            self.take(4, field)?.try_into().unwrap_or_else(|error| {
                crate::error::panic_error_invariant(
                    "operation was checked before extraction",
                    error,
                )
            }),
        ))
    }

    fn read_u64(&mut self, field: &str) -> NativeResult<u64> {
        Ok(u64::from_le_bytes(
            self.take(8, field)?.try_into().unwrap_or_else(|error| {
                crate::error::panic_error_invariant(
                    "operation was checked before extraction",
                    error,
                )
            }),
        ))
    }

    fn read_i64(&mut self, field: &str) -> NativeResult<i64> {
        Ok(i64::from_le_bytes(
            self.take(8, field)?.try_into().unwrap_or_else(|error| {
                crate::error::panic_error_invariant(
                    "operation was checked before extraction",
                    error,
                )
            }),
        ))
    }

    fn rest(&self) -> &'a [u8] {
        &self.bytes[self.offset..]
    }
}

fn usize_from_u64(value: u64, field: &str) -> NativeResult<usize> {
    usize::try_from(value).map_err(|_source| NativeError::new(format!("{field} exceeds usize")))
}

fn usize_from_i64(value: i64, field: &str) -> NativeResult<usize> {
    usize::try_from(value).map_err(|_source| NativeError::new(format!("invalid {field}")))
}

fn bound_count(value: u64, maximum: usize, field: &str) -> NativeResult<usize> {
    let value = usize_from_u64(value, field)?;
    if value > maximum {
        return Err(NativeError::new(format!("{field} exceeds parser bound")));
    }
    Ok(value)
}

fn bounded_product(count: u64, size: u64, field: &str) -> NativeResult<usize> {
    bound_count(count, MAX_NATIVE_ITEMS, field)?;
    usize_from_u64(
        count
            .checked_mul(size)
            .ok_or_else(|| NativeError::new(format!("{field} size overflow")))?,
        field,
    )
}

fn require_zeroes(bytes: &[u8], field: &str) -> NativeResult<()> {
    if bytes.iter().any(|byte| *byte != 0) {
        return Err(NativeError::new(format!("nonzero {field}")));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn push_hash_header(bytes: &mut Vec<u8>, algorithm: i32, bins: i64) {
        bytes.extend_from_slice(&algorithm.to_le_bytes());
        bytes.extend_from_slice(&8u32.to_le_bytes());
        bytes.extend_from_slice(&64u32.to_le_bytes());
        bytes.extend_from_slice(&6u32.to_le_bytes());
        bytes.extend_from_slice(&bins.to_le_bytes());
    }

    fn valid_string_dictionary(compressed: bool) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&2i32.to_le_bytes());
        bytes.extend_from_slice(&1u64.to_le_bytes());
        bytes.push(u8::from(compressed));
        bytes.extend_from_slice(&1u64.to_le_bytes());
        bytes.extend_from_slice(&1u64.to_le_bytes());
        bytes.extend_from_slice(&u64::from(compressed).to_le_bytes());
        bytes.push(0);
        bytes.extend_from_slice(&0u64.to_le_bytes());
        bytes.extend_from_slice(&1u64.to_le_bytes());
        bytes.push(u8::from(compressed));
        bytes.extend_from_slice(&STRING_PAGE_BEGIN_MARK.to_le_bytes());
        if compressed {
            bytes.extend_from_slice(&1u32.to_le_bytes());
            bytes.extend_from_slice(&703_122u32.to_le_bytes());
            bytes.extend_from_slice(&1u64.to_le_bytes());
            bytes.extend_from_slice(&2u32.to_le_bytes());
            bytes.extend_from_slice(&[0u8; 128]);
            bytes.extend_from_slice(&1u64.to_le_bytes());
            bytes.push(0);
        } else {
            bytes.extend_from_slice(&0u64.to_le_bytes());
            bytes.extend_from_slice(&1u64.to_le_bytes());
            bytes.extend_from_slice(&2u64.to_le_bytes());
            bytes.extend_from_slice(&[b'a', 0]);
        }
        bytes.extend_from_slice(&STRING_PAGE_END_MARK.to_le_bytes());
        bytes.extend_from_slice(&1u64.to_le_bytes());
        bytes.extend_from_slice(&8u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes
    }

    fn valid_hash_index() -> Vec<u8> {
        let mut bytes = Vec::new();
        push_hash_header(&mut bytes, -1, 16);
        bytes.extend_from_slice(&1u64.to_le_bytes());
        bytes.extend_from_slice(&15u64.to_le_bytes());
        bytes.push(0);
        for bin in 0..16usize {
            let mut raw = [0u8; 64];
            if bin == 0 {
                raw[8..12].copy_from_slice(&1u32.to_le_bytes());
                raw[12..16].copy_from_slice(&0u32.to_le_bytes());
                raw[16..20].copy_from_slice(&0u32.to_le_bytes());
            }
            bytes.extend_from_slice(&raw);
        }
        bytes.extend_from_slice(&0u64.to_le_bytes());
        bytes
    }

    #[test]
    fn parses_borrowed_idf_numeric_and_string_layouts() {
        let mut idf = Vec::new();
        idf.extend_from_slice(&1u64.to_le_bytes());
        idf.extend_from_slice(&[7u8; 8]);
        idf.extend_from_slice(&0u64.to_le_bytes());
        let parsed = parse_idf(&idf).unwrap();
        assert_eq!(parsed.segments.len(), 2);
        assert!(std::ptr::eq(
            parsed.segments[0].bytes.as_ptr(),
            idf[8..].as_ptr()
        ));

        let mut numeric = Vec::new();
        numeric.extend_from_slice(&0i32.to_le_bytes());
        push_hash_header(&mut numeric, -1, -1);
        numeric.extend_from_slice(&2u64.to_le_bytes());
        numeric.extend_from_slice(&4u32.to_le_bytes());
        numeric.extend_from_slice(&1i32.to_le_bytes());
        numeric.extend_from_slice(&2i32.to_le_bytes());
        let parsed = parse_dictionary(&numeric, StringHashMode::Auto).unwrap();
        assert_eq!(parsed.dictionary_type, DictionaryType::Long);

        for compressed in [false, true] {
            let bytes = valid_string_dictionary(compressed);
            let parsed = parse_dictionary(&bytes, StringHashMode::Absent).unwrap();
            let DictionaryBody::String(body) = parsed.body else {
                panic!()
            };
            assert_eq!(body.pages.len(), 1);
            assert_eq!(body.record_handles[0].offset, 0);
        }
    }

    #[test]
    fn parses_and_writes_hash_index_exactly() {
        let bytes = valid_hash_index();
        let parsed = parse_hash_index(&bytes).unwrap();
        assert_eq!(parsed.record_count, 1);
        assert_eq!(parsed.bins[0].entry_count, 1);
        let file = NativeFile {
            storage_path: "1.H$T$C.hidx",
            bytes: &bytes,
            data: NativeData::HashIndex(parsed),
        };
        assert_eq!(write_file(&file).unwrap(), bytes);
    }

    #[test]
    fn rejects_truncation_counts_constants_flags_and_ranges() {
        assert!(parse_idf(&1u64.to_le_bytes()).is_err());
        let mut idf = Vec::new();
        idf.extend_from_slice(&u64::MAX.to_le_bytes());
        assert!(parse_idf(&idf).is_err());

        let mut string = valid_string_dictionary(false);
        string[4 + 8] = 2;
        assert!(parse_dictionary(&string, StringHashMode::Absent).is_err());
        let mut string = valid_string_dictionary(false);
        let begin_mark = 4 + 8 + 1 + 8 + 8 + 8 + 1 + 8 + 8 + 1;
        string[begin_mark] ^= 1;
        assert!(parse_dictionary(&string, StringHashMode::Absent).is_err());

        let mut hash = valid_hash_index();
        hash[24 + 8] ^= 1;
        assert!(parse_hash_index(&hash).is_err());
        let mut hash = valid_hash_index();
        hash[24 + 8 + 8] = 2;
        assert!(parse_hash_index(&hash).is_err());

        let mut uncompressed = valid_string_dictionary(false);
        uncompressed[75..83]
            .copy_from_slice(&(MAX_UNCOMPRESSED_STRING_PAGE_BYTES + 1).to_le_bytes());
        assert!(
            parse_dictionary(&uncompressed, StringHashMode::Absent)
                .unwrap_err()
                .to_string()
                .contains("byte limit")
        );

        let mut compressed = valid_string_dictionary(true);
        compressed[67..75].copy_from_slice(&(MAX_COMPRESSED_STRING_PAGE_BYTES + 1).to_le_bytes());
        assert!(
            parse_dictionary(&compressed, StringHashMode::Absent)
                .unwrap_err()
                .to_string()
                .contains("byte limit")
        );
    }

    #[test]
    fn enforces_column_data_hybrid_pairs_at_the_storage_role_boundary() {
        let generated =
            classify_generated_path("Model.1.db/Table.0.dim/1.Table.Col.0.idf").unwrap();
        assert_eq!(generated.kind, GeneratedNameKind::ColumnData);

        let one_segment = 0u64.to_le_bytes();
        assert!(parse_storage_idf(generated.kind, &one_segment).is_err());

        let two_segments = [0u8; 16];
        assert_eq!(
            parse_storage_idf(generated.kind, &two_segments)
                .unwrap()
                .segments
                .len(),
            2
        );

        let three_segments = [0u8; 24];
        assert!(parse_storage_idf(generated.kind, &three_segments).is_err());
    }

    #[test]
    fn rejects_hash_misplacement_collisions_and_statistics() {
        let mut hash = valid_hash_index();
        let bins_offset = 24 + 8 + 8 + 1;
        hash[bins_offset + 12] = 1;
        assert!(parse_hash_index(&hash).is_err());

        let mut hash = valid_hash_index();
        let collision_count = hash.len() - 8;
        hash[collision_count] = 1;
        assert!(parse_hash_index(&hash).is_err());

        let mut hash = valid_hash_index();
        hash[24 + 8 + 8] = 1;
        hash.splice(24 + 8 + 8 + 1..24 + 8 + 8 + 1, [0u8; 60]);
        assert!(parse_hash_index(&hash).is_err());
    }
}
