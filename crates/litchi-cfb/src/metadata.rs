//! Strict parsing and projection of OLE Property Set streams.

use super::consts::*;
use super::file::{OleError, OleFile};
use super::writer::OleWriter;
use chrono::{DateTime, Duration, Utc};
use litchi_codepage::Mbcs;
use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Seek};

const PROPERTY_SET_HEADER_SIZE: usize = 28;
const SECTION_DESCRIPTOR_SIZE: usize = 20;
const SECTION_HEADER_SIZE: usize = 8;
const PROPERTY_DESCRIPTOR_SIZE: usize = 8;
const MAX_PROPERTY_COUNT: usize = 16_384;
const MAX_VECTOR_ELEMENTS: usize = 1_000_000;
const DEFAULT_CODEPAGE: u16 = 1252;
const UNICODE_CODEPAGE: u16 = 1200;
const PID_DICTIONARY: u32 = 0;
const PID_CODEPAGE: u32 = 1;

/// A serialized OLE GUID. The byte representation is preserved exactly as stored.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PropertySetGuid([u8; 16]);

impl PropertySetGuid {
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    pub const fn as_bytes(&self) -> &[u8; 16] {
        &self.0
    }
}

/// Text encoding permitted by an OLE Property Set section.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CodePage {
    /// UTF-16 little endian (code page 1200).
    Utf16Le,
    /// Checked NUL-terminated byte-stream code page.
    Mbcs(Mbcs),
}

impl CodePage {
    /// Windows-1252, the Property Set default.
    pub const WINDOWS_1252: Self = Self::Mbcs(Mbcs::WINDOWS_1252);

    /// Validate a raw Property Set code-page identifier.
    pub fn new(page: u16) -> Option<Self> {
        if page == UNICODE_CODEPAGE {
            Some(Self::Utf16Le)
        } else {
            Mbcs::new(u32::from(page)).map(Self::Mbcs)
        }
    }

    /// Numeric identifier stored in PID 1.
    pub const fn id(self) -> u16 {
        match self {
            Self::Utf16Le => UNICODE_CODEPAGE,
            Self::Mbcs(page) => page.id16(),
        }
    }
}

impl TryFrom<u16> for CodePage {
    type Error = OleError;

    fn try_from(page: u16) -> Result<Self, Self::Error> {
        Self::new(page).ok_or_else(|| invalid(format!("unsupported Property Set code page {page}")))
    }
}

/// One complete OLE Property Set stream.
#[derive(Debug, Clone, PartialEq)]
pub struct PropertySetStream {
    pub version: u16,
    pub system_identifier: u32,
    pub class_identifier: PropertySetGuid,
    pub sections: Vec<PropertySet>,
}

/// One independently bounded section within a Property Set stream.
#[derive(Debug, Clone, PartialEq)]
pub struct PropertySet {
    pub format_identifier: PropertySetGuid,
    codepage: Option<CodePage>,
    dictionary: HashMap<u32, String>,
    properties: HashMap<u32, PropertyValue>,
    /// Property descriptor order as stored on disk, including PID 0 when present.
    property_order: Vec<u32>,
    /// Dictionary entry order as stored on disk.
    dictionary_order: Vec<u32>,
}

impl PropertySet {
    pub fn new(format_identifier: PropertySetGuid) -> Self {
        Self {
            format_identifier,
            codepage: None,
            dictionary: HashMap::new(),
            properties: HashMap::new(),
            property_order: Vec::new(),
            dictionary_order: Vec::new(),
        }
    }
    pub fn property(&self, identifier: u32) -> Option<&PropertyValue> {
        self.properties.get(&identifier)
    }

    pub fn property_name(&self, identifier: u32) -> Option<&str> {
        self.dictionary.get(&identifier).map(String::as_str)
    }

    pub fn named_properties(&self) -> impl Iterator<Item = (&str, &PropertyValue)> {
        self.dictionary_order
            .iter()
            .filter_map(|identifier| {
                self.dictionary
                    .get(identifier)
                    .map(|name| (*identifier, name))
            })
            .filter_map(|(identifier, name)| {
                self.properties
                    .get(&identifier)
                    .map(|value| (name.as_str(), value))
            })
    }

    pub fn property_ids(&self) -> impl Iterator<Item = u32> + '_ {
        self.property_order.iter().copied().filter(|identifier| {
            *identifier != PID_DICTIONARY && self.properties.contains_key(identifier)
        })
    }

    pub fn find_named(&self, name: &str) -> Option<(u32, &PropertyValue)> {
        self.dictionary_order.iter().find_map(|identifier| {
            self.dictionary
                .get(identifier)
                .filter(|candidate| candidate.eq_ignore_ascii_case(name))
                .and_then(|_| {
                    self.properties
                        .get(identifier)
                        .map(|value| (*identifier, value))
                })
        })
    }

    pub fn add(&mut self, identifier: u32, value: PropertyValue) -> Result<(), OleError> {
        if matches!(identifier, PID_DICTIONARY | PID_CODEPAGE)
            || self.properties.contains_key(&identifier)
        {
            return Err(invalid(format!(
                "Duplicate or reserved property identifier {identifier}"
            )));
        }
        self.properties.insert(identifier, value);
        self.property_order.push(identifier);
        Ok(())
    }

    pub fn add_named(
        &mut self,
        identifier: u32,
        name: String,
        value: PropertyValue,
    ) -> Result<(), OleError> {
        validate_property_name(&name)?;
        if self
            .dictionary
            .values()
            .any(|existing| existing.eq_ignore_ascii_case(&name))
        {
            return Err(invalid(format!("Duplicate property name '{name}'")));
        }
        self.add(identifier, value)?;
        self.dictionary.insert(identifier, name);
        self.dictionary_order.push(identifier);
        if !self.property_order.contains(&PID_DICTIONARY) {
            self.property_order.insert(0, PID_DICTIONARY);
        }
        Ok(())
    }

    pub fn update(
        &mut self,
        identifier: u32,
        value: PropertyValue,
    ) -> Result<PropertyValue, OleError> {
        if identifier == PID_CODEPAGE {
            return Err(invalid("Use set_page to update PID 1"));
        }
        let target = self
            .properties
            .get_mut(&identifier)
            .ok_or_else(|| invalid(format!("Property {identifier} does not exist")))?;
        Ok(std::mem::replace(target, value))
    }

    pub fn replace(&mut self, identifier: u32, value: PropertyValue) -> Option<PropertyValue> {
        if matches!(identifier, PID_DICTIONARY | PID_CODEPAGE) {
            return None;
        }
        if !self.properties.contains_key(&identifier) {
            self.property_order.push(identifier);
        }
        self.properties.insert(identifier, value)
    }

    pub fn remove(&mut self, identifier: u32) -> Option<PropertyValue> {
        if identifier == PID_CODEPAGE {
            return None;
        }
        self.property_order.retain(|value| *value != identifier);
        self.dictionary_order.retain(|value| *value != identifier);
        self.dictionary.remove(&identifier);
        let removed = self.properties.remove(&identifier);
        if self.dictionary.is_empty() {
            self.property_order.retain(|value| *value != PID_DICTIONARY);
        }
        removed
    }

    pub fn remove_named(&mut self, name: &str) -> Option<PropertyValue> {
        let identifier = self.dictionary_order.iter().copied().find(|identifier| {
            self.dictionary
                .get(identifier)
                .is_some_and(|value| value.eq_ignore_ascii_case(name))
        })?;
        self.remove(identifier)
    }

    pub fn rename(&mut self, identifier: u32, name: String) -> Result<(), OleError> {
        if identifier == PID_CODEPAGE {
            return Err(invalid("PID 1 cannot be a named property"));
        }
        validate_property_name(&name)?;
        if self
            .dictionary
            .iter()
            .any(|(other, existing)| *other != identifier && existing.eq_ignore_ascii_case(&name))
        {
            return Err(invalid(format!("Duplicate property name '{name}'")));
        }
        if !self.properties.contains_key(&identifier) {
            return Err(invalid(format!("Property {identifier} does not exist")));
        }
        self.dictionary.insert(identifier, name);
        if !self.dictionary_order.contains(&identifier) {
            self.dictionary_order.push(identifier);
        }
        if !self.property_order.contains(&PID_DICTIONARY) {
            self.property_order.insert(0, PID_DICTIONARY);
        }
        Ok(())
    }

    pub fn reorder(&mut self, order: &[u32]) -> Result<(), OleError> {
        let current = self.property_order.iter().copied().collect::<HashSet<_>>();
        let proposed = order.iter().copied().collect::<HashSet<_>>();
        if current != proposed || proposed.len() != order.len() {
            return Err(invalid(
                "Property reorder must contain every identifier exactly once",
            ));
        }
        self.property_order = order.to_vec();
        Ok(())
    }

    pub fn clear(&mut self) {
        self.properties.clear();
        self.dictionary.clear();
        self.property_order.clear();
        self.dictionary_order.clear();
        self.codepage = None;
    }

    /// Return the checked text page declared by PID 1.
    pub const fn page(&self) -> Option<CodePage> {
        self.codepage
    }

    /// Set PID 1 from a checked text-page capability.
    pub fn set_page(&mut self, page: CodePage) {
        let codepage = page.id();
        self.codepage = Some(page);
        self.properties
            .insert(PID_CODEPAGE, PropertyValue::I2(codepage as i16));
        if !self.property_order.contains(&PID_CODEPAGE) {
            self.property_order.insert(0, PID_CODEPAGE);
        }
    }

    /// Validate a raw identifier and set PID 1.
    pub fn set_page_id(&mut self, page: u16) -> Result<(), OleError> {
        self.set_page(CodePage::try_from(page)?);
        Ok(())
    }

    /// Remove the typed PID 1 declaration.
    pub fn clear_page(&mut self) -> Option<CodePage> {
        self.property_order.retain(|value| *value != PID_CODEPAGE);
        self.properties.remove(&PID_CODEPAGE);
        self.codepage.take()
    }
}

/// Metadata projected from SummaryInformation and DocumentSummaryInformation.
#[derive(Debug, Clone, Default)]
pub struct OleMetadata {
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
    pub category: Option<String>,
    pub manager: Option<String>,
    pub company: Option<String>,
    pub custom_properties: HashMap<String, PropertyValue>,
}

/// A typed OLE property value. Unsupported variants retain their bounded raw bytes.
#[derive(Debug, Clone, PartialEq)]
pub enum PropertyValue {
    Empty,
    Null,
    I1(i8),
    UI1(u8),
    I2(i16),
    UI2(u16),
    I4(i32),
    UI4(u32),
    I8(i64),
    UI8(u64),
    Int(i32),
    UInt(u32),
    R4(f32),
    R8(f64),
    Currency(i64),
    Date(f64),
    Bstr(String),
    Error(u32),
    Bool(bool),
    Decimal([u8; 16]),
    Lpstr(String),
    Lpwstr(String),
    Filetime(u64),
    Blob(Vec<u8>),
    Clipboard { format: i32, data: Vec<u8> },
    Clsid(PropertySetGuid),
    Vector(Vec<PropertyValue>),
    Unknown { variant_type: u16, data: Vec<u8> },
}

impl PropertySetStream {
    pub fn new(section: PropertySet) -> Self {
        Self {
            version: 0,
            system_identifier: 0,
            class_identifier: PropertySetGuid::from_bytes([0; 16]),
            sections: vec![section],
        }
    }
    /// Parse a complete Property Set stream with section-local bounds.
    pub fn parse(data: &[u8]) -> Result<Self, OleError> {
        if data.len() < PROPERTY_SET_HEADER_SIZE + SECTION_DESCRIPTOR_SIZE {
            return Err(invalid("Property Set stream is too short"));
        }
        if read_u16(data, 0, "byte order")? != 0xfffe {
            return Err(invalid("Property Set byte order must be 0xFFFE"));
        }
        let version = read_u16(data, 2, "version")?;
        if version != 0 {
            return Err(invalid(format!(
                "Unsupported Property Set version {version}"
            )));
        }
        let system_identifier = read_u32(data, 4, "system identifier")?;
        let class_identifier = read_guid(data, 8, "class identifier")?;
        let section_count = usize::try_from(read_u32(data, 24, "section count")?)
            .map_err(|_| invalid("Property Set section count is too large"))?;
        if !(1..=2).contains(&section_count) {
            return Err(invalid(format!(
                "Property Set section count {section_count} is outside 1..=2"
            )));
        }
        let descriptor_end = section_count
            .checked_mul(SECTION_DESCRIPTOR_SIZE)
            .and_then(|size| PROPERTY_SET_HEADER_SIZE.checked_add(size))
            .ok_or_else(|| invalid("Property Set descriptor table overflow"))?;
        checked_range(data, 0, descriptor_end, "descriptor table")?;

        let mut descriptors = Vec::with_capacity(section_count);
        let mut format_ids = HashSet::with_capacity(section_count);
        let mut offsets = HashSet::with_capacity(section_count);
        for index in 0..section_count {
            let descriptor = PROPERTY_SET_HEADER_SIZE + index * SECTION_DESCRIPTOR_SIZE;
            let format_identifier = read_guid(data, descriptor, "section format identifier")?;
            let offset = usize::try_from(read_u32(data, descriptor + 16, "section offset")?)
                .map_err(|_| invalid("Property Set section offset is too large"))?;
            if offset < descriptor_end || offset % 4 != 0 || offset >= data.len() {
                return Err(invalid(format!(
                    "Invalid Property Set section offset {offset}"
                )));
            }
            if !format_ids.insert(format_identifier) || !offsets.insert(offset) {
                return Err(invalid("Duplicate Property Set section descriptor"));
            }
            let size = usize::try_from(read_u32(data, offset, "section size")?)
                .map_err(|_| invalid("Property Set section size is too large"))?;
            let end = offset
                .checked_add(size)
                .filter(|end| *end <= data.len())
                .ok_or_else(|| invalid("Property Set section exceeds its stream"))?;
            descriptors.push((format_identifier, offset, end));
        }

        let mut ranges: Vec<(usize, usize)> = descriptors
            .iter()
            .map(|(_, start, end)| (*start, *end))
            .collect();
        ranges.sort_unstable();
        if ranges.windows(2).any(|pair| pair[1].0 < pair[0].1) {
            return Err(invalid("Property Set sections overlap"));
        }

        let mut sections = Vec::with_capacity(section_count);
        for (format_identifier, start, end) in descriptors {
            sections.push(parse_section(&data[start..end], format_identifier)?);
        }
        Ok(Self {
            version,
            system_identifier,
            class_identifier,
            sections,
        })
    }

    pub fn section(&self, format_identifier: PropertySetGuid) -> Option<&PropertySet> {
        self.sections
            .iter()
            .find(|section| section.format_identifier == format_identifier)
    }
    pub fn section_mut(&mut self, format_identifier: PropertySetGuid) -> Option<&mut PropertySet> {
        self.sections
            .iter_mut()
            .find(|section| section.format_identifier == format_identifier)
    }
    pub fn add_section(&mut self, section: PropertySet) -> Result<(), OleError> {
        if self.sections.len() == 2 || self.section(section.format_identifier).is_some() {
            return Err(invalid("Duplicate or excess Property Set section"));
        }
        self.sections.push(section);
        Ok(())
    }
    pub fn remove_section(&mut self, format_identifier: PropertySetGuid) -> Option<PropertySet> {
        let index = self
            .sections
            .iter()
            .position(|section| section.format_identifier == format_identifier)?;
        Some(self.sections.remove(index))
    }
    pub fn reorder_sections(&mut self, order: &[PropertySetGuid]) -> Result<(), OleError> {
        if order.len() != self.sections.len()
            || order.iter().copied().collect::<HashSet<_>>().len() != order.len()
        {
            return Err(invalid("Section reorder is incomplete or duplicated"));
        }
        let mut reordered = Vec::with_capacity(order.len());
        for id in order {
            let index = self
                .sections
                .iter()
                .position(|section| section.format_identifier == *id)
                .ok_or_else(|| invalid("Section reorder references an unknown format ID"))?;
            reordered.push(self.sections[index].clone())
        }
        self.sections = reordered;
        Ok(())
    }
    pub fn clear_sections(&mut self) {
        self.sections.clear();
    }
    pub fn to_bytes(&self) -> Result<Vec<u8>, OleError> {
        serialize_property_set_stream(self)
    }
}

pub const SUMMARY_INFORMATION_FMTID: PropertySetGuid = PropertySetGuid::from_bytes([
    0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9,
]);
pub const DOCUMENT_SUMMARY_INFORMATION_FMTID: PropertySetGuid = PropertySetGuid::from_bytes([
    0x02, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE,
]);
pub const USER_DEFINED_PROPERTIES_FMTID: PropertySetGuid = PropertySetGuid::from_bytes([
    0x05, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE,
]);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StandardPropertySet {
    SummaryInformation,
    DocumentSummaryInformation,
    UserDefinedProperties,
}

impl StandardPropertySet {
    fn path(self) -> &'static str {
        match self {
            Self::SummaryInformation => "\u{0005}SummaryInformation",
            Self::DocumentSummaryInformation | Self::UserDefinedProperties => {
                "\u{0005}DocumentSummaryInformation"
            },
        }
    }
    fn format_id(self) -> PropertySetGuid {
        match self {
            Self::SummaryInformation => SUMMARY_INFORMATION_FMTID,
            Self::DocumentSummaryInformation => DOCUMENT_SUMMARY_INFORMATION_FMTID,
            Self::UserDefinedProperties => USER_DEFINED_PROPERTIES_FMTID,
        }
    }
}

/// Transactional editor for standard OLE Property Set streams.
pub struct OlePropertySetEditor {
    original: Vec<u8>,
    streams: Vec<(Vec<String>, Vec<u8>)>,
    staged: HashMap<String, Option<Vec<u8>>>,
}

impl OlePropertySetEditor {
    pub fn new(bytes: Vec<u8>) -> Result<Self, OleError> {
        let mut ole = OleFile::open(Cursor::new(bytes.clone()))?;
        let paths = ole.list_streams();
        if paths.iter().flatten().any(|name| {
            matches!(
                name.to_ascii_lowercase().as_str(),
                "encryptedpackage" | "encryptioninfo"
            ) || name.to_ascii_lowercase().contains("digitalsignature")
        }) {
            return Err(invalid(
                "Property Set mutation is not permitted on encrypted or signed containers",
            ));
        }
        let mut streams = Vec::with_capacity(paths.len());
        for path in paths {
            let data = {
                let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
                ole.open_stream(&refs)?
            };
            streams.push((path, data));
        }
        Ok(Self {
            original: bytes,
            streams,
            staged: HashMap::new(),
        })
    }

    pub fn property_set(&self, kind: StandardPropertySet) -> Result<Option<PropertySet>, OleError> {
        let Some(stream) = self.load_stream(kind)? else {
            return Ok(None);
        };
        Ok(stream.section(kind.format_id()).cloned())
    }

    pub fn update<F>(&mut self, kind: StandardPropertySet, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut PropertySet) -> Result<(), OleError>,
    {
        let mut stream = self.load_stream(kind)?.unwrap_or_else(|| {
            let base = if kind == StandardPropertySet::UserDefinedProperties {
                DOCUMENT_SUMMARY_INFORMATION_FMTID
            } else {
                kind.format_id()
            };
            PropertySetStream::new(PropertySet::new(base))
        });
        if stream.section(kind.format_id()).is_none() {
            stream.add_section(PropertySet::new(kind.format_id()))?;
        }
        let section = stream
            .section_mut(kind.format_id())
            .expect("section inserted");
        let mut candidate = section.clone();
        edit(&mut candidate)?;
        validate_section(&candidate)?;
        *section = candidate;
        let bytes = stream.to_bytes()?;
        self.staged.insert(kind.path().to_string(), Some(bytes));
        Ok(())
    }

    pub fn replace(
        &mut self,
        kind: StandardPropertySet,
        section: PropertySet,
    ) -> Result<Option<PropertySet>, OleError> {
        if section.format_identifier != kind.format_id() {
            return Err(invalid(
                "Replacement section format ID does not match target",
            ));
        }
        let previous = self.property_set(kind)?;
        self.update(kind, |target| {
            *target = section;
            Ok(())
        })?;
        Ok(previous)
    }

    pub fn remove(&mut self, kind: StandardPropertySet) -> Result<Option<PropertySet>, OleError> {
        let Some(previous) = self.property_set(kind)? else {
            return Ok(None);
        };
        if kind == StandardPropertySet::UserDefinedProperties {
            let mut stream = self.load_stream(kind)?.expect("property existed");
            stream.remove_section(USER_DEFINED_PROPERTIES_FMTID);
            self.staged
                .insert(kind.path().to_string(), Some(stream.to_bytes()?));
        } else {
            self.staged.insert(kind.path().to_string(), None);
        }
        Ok(Some(previous))
    }

    pub fn finish(self) -> Result<Vec<u8>, OleError> {
        if self.staged.is_empty() {
            return Ok(self.original);
        }
        let mut writer = OleWriter::new();
        let mut written = HashSet::new();
        for (path, data) in self.streams {
            let key = path.last().cloned().unwrap_or_default();
            let replacement = self.staged.get(&key);
            if replacement.is_some_and(Option::is_none) {
                continue;
            }
            let bytes = replacement.and_then(Option::as_ref).unwrap_or(&data);
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            writer.create_stream(&refs, bytes)?;
            written.insert(key);
        }
        for (path, replacement) in &self.staged {
            if !written.contains(path)
                && let Some(data) = replacement
            {
                writer.create_stream(&[path], data)?;
            }
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }

    fn load_stream(
        &self,
        kind: StandardPropertySet,
    ) -> Result<Option<PropertySetStream>, OleError> {
        let path = kind.path();
        if let Some(staged) = self.staged.get(path) {
            return staged
                .as_ref()
                .map(|data| PropertySetStream::parse(data))
                .transpose();
        }
        self.streams
            .iter()
            .find(|(candidate, _)| candidate.len() == 1 && candidate[0] == path)
            .map(|(_, data)| PropertySetStream::parse(data))
            .transpose()
    }
}

fn validate_property_name(name: &str) -> Result<(), OleError> {
    if name.is_empty() || name.chars().any(|value| value == '\0') {
        Err(invalid("Property names must be nonempty and NUL-free"))
    } else {
        Ok(())
    }
}

fn validate_section(section: &PropertySet) -> Result<(), OleError> {
    if section.properties.len() > MAX_PROPERTY_COUNT {
        return Err(invalid("Property count exceeds safety limit"));
    }
    if section
        .dictionary
        .values()
        .collect::<Vec<_>>()
        .iter()
        .enumerate()
        .any(|(index, name)| {
            section
                .dictionary
                .values()
                .skip(index + 1)
                .any(|other| name.eq_ignore_ascii_case(other))
        })
    {
        return Err(invalid("Duplicate dictionary property name"));
    }
    for (identifier, name) in &section.dictionary {
        validate_property_name(name)?;
        if !section.properties.contains_key(identifier) {
            return Err(invalid("Dictionary references a missing property"));
        }
    }
    match (section.codepage, section.properties.get(&PID_CODEPAGE)) {
        (Some(codepage), Some(PropertyValue::I2(value))) if *value == codepage.id() as i16 => {},
        (None, None) => {},
        _ => return Err(invalid("PID 1 does not match section codepage")),
    }
    Ok(())
}

fn serialize_property_set_stream(stream: &PropertySetStream) -> Result<Vec<u8>, OleError> {
    if stream.version != 0 || !(1..=2).contains(&stream.sections.len()) {
        return Err(invalid(
            "Property Set must have version zero and one or two sections",
        ));
    }
    let mut ids = HashSet::new();
    for section in &stream.sections {
        if !ids.insert(section.format_identifier) {
            return Err(invalid("Duplicate section format identifier"));
        }
        validate_section(section)?
    }
    let descriptor_end = PROPERTY_SET_HEADER_SIZE + stream.sections.len() * SECTION_DESCRIPTOR_SIZE;
    let mut sections = Vec::new();
    for section in &stream.sections {
        sections.push(serialize_section(section)?)
    }
    let mut offsets = Vec::new();
    let mut cursor = align4_len(descriptor_end);
    for section in &sections {
        offsets.push(cursor);
        cursor = align4_len(
            cursor
                .checked_add(section.len())
                .ok_or_else(|| invalid("Property Set size overflow"))?,
        )
    }
    let mut out = vec![0u8; align4_len(descriptor_end)];
    out[0..2].copy_from_slice(&0xfffeu16.to_le_bytes());
    out[2..4].copy_from_slice(&stream.version.to_le_bytes());
    out[4..8].copy_from_slice(&stream.system_identifier.to_le_bytes());
    out[8..24].copy_from_slice(stream.class_identifier.as_bytes());
    out[24..28].copy_from_slice(&(stream.sections.len() as u32).to_le_bytes());
    for (index, section) in stream.sections.iter().enumerate() {
        let base = PROPERTY_SET_HEADER_SIZE + index * SECTION_DESCRIPTOR_SIZE;
        out[base..base + 16].copy_from_slice(section.format_identifier.as_bytes());
        out[base + 16..base + 20].copy_from_slice(
            &u32::try_from(offsets[index])
                .map_err(|_| invalid("Section offset overflow"))?
                .to_le_bytes(),
        )
    }
    for (index, section) in sections.into_iter().enumerate() {
        out.resize(offsets[index], 0);
        out.extend_from_slice(&section);
        while !out.len().is_multiple_of(4) {
            out.push(0)
        }
    }
    Ok(out)
}

fn serialize_section(section: &PropertySet) -> Result<Vec<u8>, OleError> {
    let mut order = section.property_order.clone();
    if !section.dictionary.is_empty() && !order.contains(&PID_DICTIONARY) {
        order.insert(0, PID_DICTIONARY)
    }
    for id in section.properties.keys() {
        if !order.contains(id) {
            order.push(*id)
        }
    }
    order.retain(|id| {
        *id == PID_DICTIONARY && !section.dictionary.is_empty()
            || section.properties.contains_key(id)
    });
    if order.iter().copied().collect::<HashSet<_>>().len() != order.len() {
        return Err(invalid("Duplicate property order identifier"));
    }
    let table_end = SECTION_HEADER_SIZE + order.len() * PROPERTY_DESCRIPTOR_SIZE;
    let mut values = Vec::new();
    let codepage = section.codepage.unwrap_or(CodePage::WINDOWS_1252).id();
    for id in &order {
        values.push(if *id == PID_DICTIONARY {
            serialize_dictionary(section, codepage)?
        } else {
            let value = section
                .properties
                .get(id)
                .ok_or_else(|| invalid(format!("Property order references missing PID {id}")))?;
            serialize_typed(value, codepage)?
        })
    }
    let mut offsets = Vec::new();
    let mut cursor = align4_len(table_end);
    for value in &values {
        offsets.push(cursor);
        cursor = align4_len(
            cursor
                .checked_add(value.len())
                .ok_or_else(|| invalid("Section size overflow"))?,
        )
    }
    let mut out = vec![0u8; align4_len(table_end)];
    out[4..8].copy_from_slice(&(order.len() as u32).to_le_bytes());
    for (index, id) in order.iter().enumerate() {
        let base = SECTION_HEADER_SIZE + index * PROPERTY_DESCRIPTOR_SIZE;
        out[base..base + 4].copy_from_slice(&id.to_le_bytes());
        out[base + 4..base + 8].copy_from_slice(
            &u32::try_from(offsets[index])
                .map_err(|_| invalid("Property offset overflow"))?
                .to_le_bytes(),
        )
    }
    for (index, value) in values.into_iter().enumerate() {
        out.resize(offsets[index], 0);
        out.extend_from_slice(&value);
        while !out.len().is_multiple_of(4) {
            out.push(0)
        }
    }
    let section_len = u32::try_from(out.len()).map_err(|_| invalid("Section length overflow"))?;
    out[0..4].copy_from_slice(&section_len.to_le_bytes());
    Ok(out)
}

fn serialize_dictionary(section: &PropertySet, codepage: u16) -> Result<Vec<u8>, OleError> {
    let mut order = section.dictionary_order.clone();
    for id in section.dictionary.keys() {
        if !order.contains(id) {
            order.push(*id)
        }
    }
    order.retain(|id| section.dictionary.contains_key(id));
    if order.iter().copied().collect::<HashSet<_>>().len() != order.len() {
        return Err(invalid("Duplicate dictionary order identifier"));
    }
    let mut out = (order.len() as u32).to_le_bytes().to_vec();
    for id in order {
        let name = &section.dictionary[&id];
        out.extend_from_slice(&id.to_le_bytes());
        if codepage == UNICODE_CODEPAGE {
            let units = name.encode_utf16().collect::<Vec<_>>();
            out.extend_from_slice(
                &u32::try_from(units.len() + 1)
                    .map_err(|_| invalid("Dictionary name too long"))?
                    .to_le_bytes(),
            );
            for unit in units {
                out.extend_from_slice(&unit.to_le_bytes())
            }
            out.extend_from_slice(&0u16.to_le_bytes());
            while !out.len().is_multiple_of(4) {
                out.push(0)
            }
        } else {
            let bytes = encode_ansi(name, codepage)?;
            out.extend_from_slice(
                &u32::try_from(bytes.len() + 1)
                    .map_err(|_| invalid("Dictionary name too long"))?
                    .to_le_bytes(),
            );
            out.extend_from_slice(&bytes);
            out.push(0)
        }
    }
    Ok(out)
}

fn serialize_typed(value: &PropertyValue, codepage: u16) -> Result<Vec<u8>, OleError> {
    let mut out = Vec::new();
    append_typed(&mut out, value, codepage)?;
    Ok(out)
}
fn append_typed(out: &mut Vec<u8>, value: &PropertyValue, codepage: u16) -> Result<(), OleError> {
    let vt = variant_type(value);
    out.extend_from_slice(&vt.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    append_body(out, value, codepage)?;
    Ok(())
}
fn variant_type(value: &PropertyValue) -> u16 {
    match value {
        PropertyValue::Empty => VT_EMPTY,
        PropertyValue::Null => VT_NULL,
        PropertyValue::I1(_) => VT_I1,
        PropertyValue::UI1(_) => VT_UI1,
        PropertyValue::I2(_) => VT_I2,
        PropertyValue::UI2(_) => VT_UI2,
        PropertyValue::I4(_) => VT_I4,
        PropertyValue::UI4(_) => VT_UI4,
        PropertyValue::I8(_) => VT_I8,
        PropertyValue::UI8(_) => VT_UI8,
        PropertyValue::Int(_) => VT_INT,
        PropertyValue::UInt(_) => VT_UINT,
        PropertyValue::R4(_) => VT_R4,
        PropertyValue::R8(_) => VT_R8,
        PropertyValue::Currency(_) => VT_CY,
        PropertyValue::Date(_) => VT_DATE,
        PropertyValue::Bstr(_) => VT_BSTR,
        PropertyValue::Error(_) => VT_ERROR,
        PropertyValue::Bool(_) => VT_BOOL,
        PropertyValue::Decimal(_) => VT_DECIMAL,
        PropertyValue::Lpstr(_) => VT_LPSTR,
        PropertyValue::Lpwstr(_) => VT_LPWSTR,
        PropertyValue::Filetime(_) => VT_FILETIME,
        PropertyValue::Blob(_) => VT_BLOB,
        PropertyValue::Clipboard { .. } => VT_CF,
        PropertyValue::Clsid(_) => VT_CLSID,
        PropertyValue::Vector(_) => VT_VECTOR | VT_VARIANT,
        PropertyValue::Unknown { variant_type, .. } => *variant_type,
    }
}
fn append_body(out: &mut Vec<u8>, value: &PropertyValue, codepage: u16) -> Result<(), OleError> {
    match value {
        PropertyValue::Empty | PropertyValue::Null => {},
        PropertyValue::I1(v) => out.push(*v as u8),
        PropertyValue::UI1(v) => out.push(*v),
        PropertyValue::I2(v) => out.extend_from_slice(&v.to_le_bytes()),
        PropertyValue::UI2(v) => out.extend_from_slice(&v.to_le_bytes()),
        PropertyValue::I4(v) | PropertyValue::Int(v) => out.extend_from_slice(&v.to_le_bytes()),
        PropertyValue::UI4(v) | PropertyValue::UInt(v) | PropertyValue::Error(v) => {
            out.extend_from_slice(&v.to_le_bytes())
        },
        PropertyValue::I8(v) | PropertyValue::Currency(v) => {
            out.extend_from_slice(&v.to_le_bytes())
        },
        PropertyValue::UI8(v) | PropertyValue::Filetime(v) => {
            out.extend_from_slice(&v.to_le_bytes())
        },
        PropertyValue::R4(v) => out.extend_from_slice(&v.to_bits().to_le_bytes()),
        PropertyValue::R8(v) | PropertyValue::Date(v) => {
            out.extend_from_slice(&v.to_bits().to_le_bytes())
        },
        PropertyValue::Bool(v) => {
            out.extend_from_slice(&(if *v { -1i16 } else { 0 }).to_le_bytes())
        },
        PropertyValue::Decimal(v) => out.extend_from_slice(v),
        PropertyValue::Bstr(v) | PropertyValue::Lpstr(v) => {
            append_codepage_string(out, v, codepage)?
        },
        PropertyValue::Lpwstr(v) => {
            let units = v.encode_utf16().collect::<Vec<_>>();
            out.extend_from_slice(&(units.len() as u32 + 1).to_le_bytes());
            for unit in units {
                out.extend_from_slice(&unit.to_le_bytes())
            }
            out.extend_from_slice(&0u16.to_le_bytes())
        },
        PropertyValue::Blob(v) => {
            out.extend_from_slice(&(v.len() as u32).to_le_bytes());
            out.extend_from_slice(v);
            pad4(out)
        },
        PropertyValue::Clipboard { format, data } => {
            out.extend_from_slice(&(data.len() as u32 + 4).to_le_bytes());
            out.extend_from_slice(&format.to_le_bytes());
            out.extend_from_slice(data);
            pad4(out)
        },
        PropertyValue::Clsid(v) => out.extend_from_slice(v.as_bytes()),
        PropertyValue::Vector(values) => {
            if values.len() > MAX_VECTOR_ELEMENTS {
                return Err(invalid("Vector exceeds safety limit"));
            }
            out.extend_from_slice(&(values.len() as u32).to_le_bytes());
            for value in values {
                append_typed(out, value, codepage)?;
                pad4(out)
            }
        },
        PropertyValue::Unknown { data, .. } => out.extend_from_slice(data),
    }
    Ok(())
}
fn append_codepage_string(out: &mut Vec<u8>, value: &str, codepage: u16) -> Result<(), OleError> {
    let mut bytes = if codepage == UNICODE_CODEPAGE {
        value
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>()
    } else {
        encode_ansi(value, codepage)?
    };
    if codepage == UNICODE_CODEPAGE {
        bytes.extend_from_slice(&[0, 0])
    } else {
        bytes.push(0)
    }
    out.extend_from_slice(&(bytes.len() as u32).to_le_bytes());
    out.extend_from_slice(&bytes);
    pad4(out);
    Ok(())
}
fn encode_ansi(value: &str, codepage: u16) -> Result<Vec<u8>, OleError> {
    let page = Mbcs::require(u32::from(codepage)).map_err(|error| invalid(error.to_string()))?;
    page.encode(value)
        .map(Cow::into_owned)
        .map_err(|error| invalid(error.to_string()))
}
fn align4_len(value: usize) -> usize {
    (value + 3) & !3
}
fn pad4(out: &mut Vec<u8>) {
    while !out.len().is_multiple_of(4) {
        out.push(0)
    }
}

impl<R: Read + Seek> OleFile<R> {
    /// Strictly parse a Property Set stream at `path`.
    pub fn property_set_stream(&mut self, path: &[&str]) -> Result<PropertySetStream, OleError> {
        let data = self.open_stream(path)?;
        PropertySetStream::parse(&data)
    }

    /// Parse standard metadata. Missing streams are optional; malformed streams are errors.
    pub fn get_metadata(&mut self) -> Result<OleMetadata, OleError> {
        let mut metadata = OleMetadata::default();
        match self.property_set_stream(&["\u{0005}SummaryInformation"]) {
            Ok(stream) => {
                let section = stream
                    .sections
                    .first()
                    .ok_or_else(|| invalid("SummaryInformation has no section"))?;
                extract_summary_info(&mut metadata, section);
            },
            Err(OleError::StreamNotFound) => {},
            Err(error) => return Err(error),
        }
        match self.property_set_stream(&["\u{0005}DocumentSummaryInformation"]) {
            Ok(stream) => {
                let section = stream
                    .sections
                    .first()
                    .ok_or_else(|| invalid("DocumentSummaryInformation has no section"))?;
                extract_document_summary_info(&mut metadata, section);
                for custom_section in stream.sections.iter().skip(1) {
                    for (name, value) in custom_section.named_properties() {
                        if metadata
                            .custom_properties
                            .insert(name.to_string(), value.clone())
                            .is_some()
                        {
                            return Err(invalid(format!(
                                "Duplicate custom property name '{name}'"
                            )));
                        }
                    }
                }
            },
            Err(OleError::StreamNotFound) => {},
            Err(error) => return Err(error),
        }
        Ok(metadata)
    }
}

fn parse_section(data: &[u8], format_identifier: PropertySetGuid) -> Result<PropertySet, OleError> {
    checked_range(data, 0, SECTION_HEADER_SIZE, "section header")?;
    let declared_size = usize::try_from(read_u32(data, 0, "section size")?)
        .map_err(|_| invalid("Property Set section size is too large"))?;
    if declared_size != data.len() {
        return Err(invalid(
            "Property Set section range does not match its size",
        ));
    }
    let property_count = usize::try_from(read_u32(data, 4, "property count")?)
        .map_err(|_| invalid("Property count is too large"))?;
    if property_count > MAX_PROPERTY_COUNT {
        return Err(invalid(format!(
            "Property count {property_count} exceeds the safety limit"
        )));
    }
    let table_end = property_count
        .checked_mul(PROPERTY_DESCRIPTOR_SIZE)
        .and_then(|size| SECTION_HEADER_SIZE.checked_add(size))
        .ok_or_else(|| invalid("Property table size overflow"))?;
    if table_end > data.len() {
        return Err(invalid("Property table exceeds its section"));
    }

    let mut identifiers = HashSet::with_capacity(property_count);
    let mut offsets = HashSet::with_capacity(property_count);
    let mut descriptors = Vec::with_capacity(property_count);
    for index in 0..property_count {
        let descriptor = SECTION_HEADER_SIZE + index * PROPERTY_DESCRIPTOR_SIZE;
        let identifier = read_u32(data, descriptor, "property identifier")?;
        let offset = usize::try_from(read_u32(data, descriptor + 4, "property offset")?)
            .map_err(|_| invalid("Property offset is too large"))?;
        // Older Office producers emit valid properties at non-DWORD offsets.
        // Bounds and uniqueness remain mandatory, but alignment is advisory.
        if offset < table_end || offset >= data.len() {
            return Err(invalid(format!("Invalid property offset {offset}")));
        }
        if !identifiers.insert(identifier) {
            return Err(invalid(format!(
                "Duplicate property identifier {identifier}"
            )));
        }
        if !offsets.insert(offset) {
            return Err(invalid(format!("Duplicate property offset {offset}")));
        }
        descriptors.push((offset, identifier));
    }
    descriptors.sort_unstable();

    let property_slice = |identifier: u32| -> Result<Option<(usize, &[u8])>, OleError> {
        let Some(index) = descriptors.iter().position(|(_, id)| *id == identifier) else {
            return Ok(None);
        };
        let start = descriptors[index].0;
        let end = descriptors
            .get(index + 1)
            .map_or(data.len(), |(offset, _)| *offset);
        if end <= start {
            return Err(invalid("Property has an empty or inverted range"));
        }
        Ok(Some((start, &data[start..end])))
    };

    let (codepage, codepage_value) = match property_slice(PID_CODEPAGE)? {
        Some((offset, bytes)) => {
            if read_u16(bytes, 0, "codepage type")? != VT_I2
                || read_u16(bytes, 2, "codepage reserved field")? != 0
            {
                return Err(invalid("PID 1 must be a VT_I2 codepage"));
            }
            let value = parse_typed_property(bytes, DEFAULT_CODEPAGE, offset)?;
            let PropertyValue::I2(signed) = value else {
                return Err(invalid("PID 1 must contain a VT_I2 codepage"));
            };
            let codepage = CodePage::try_from(signed as u16)?;
            (Some(codepage), Some(PropertyValue::I2(signed)))
        },
        None => (None, None),
    };
    let effective_codepage = codepage.unwrap_or(CodePage::WINDOWS_1252).id();

    let (dictionary, dictionary_order) = match property_slice(PID_DICTIONARY)? {
        Some((offset, bytes)) => parse_dictionary(bytes, effective_codepage, offset)?,
        None => (HashMap::new(), Vec::new()),
    };
    let mut properties = HashMap::with_capacity(property_count.saturating_sub(1));
    if let Some(value) = codepage_value {
        properties.insert(PID_CODEPAGE, value);
    }
    for (index, (start, identifier)) in descriptors.iter().enumerate() {
        if matches!(*identifier, PID_DICTIONARY | PID_CODEPAGE) {
            continue;
        }
        let end = descriptors
            .get(index + 1)
            .map_or(data.len(), |(offset, _)| *offset);
        let bytes = &data[*start..end];
        let value = parse_typed_property(bytes, effective_codepage, *start).map_err(|error| {
            invalid(format!(
                "Property {identifier} at section offset {start} is invalid: {error}"
            ))
        })?;
        properties.insert(*identifier, value);
    }
    Ok(PropertySet {
        format_identifier,
        codepage,
        dictionary,
        properties,
        property_order: descriptors
            .iter()
            .map(|(_, identifier)| *identifier)
            .collect(),
        dictionary_order,
    })
}

fn parse_dictionary(
    data: &[u8],
    codepage: u16,
    property_offset: usize,
) -> Result<(HashMap<u32, String>, Vec<u32>), OleError> {
    let mut reader = ValueReader::new(data, property_offset);
    let count = usize::try_from(reader.read_u32("dictionary count")?)
        .map_err(|_| invalid("Dictionary count is too large"))?;
    if count > reader.remaining_len() / 8 {
        return Err(invalid("Dictionary count exceeds its property range"));
    }
    let mut dictionary = HashMap::with_capacity(count);
    let mut order = Vec::with_capacity(count);
    for _ in 0..count {
        let identifier = reader.read_u32("dictionary property identifier")?;
        let length = usize::try_from(reader.read_u32("dictionary string length")?)
            .map_err(|_| invalid("Dictionary string length is too large"))?;
        if length == 0 {
            return Err(invalid("Dictionary strings must include a terminator"));
        }
        let name = if codepage == UNICODE_CODEPAGE {
            let payload_units = length - 1;
            let payload_bytes = payload_units
                .checked_mul(2)
                .ok_or_else(|| invalid("Dictionary Unicode length overflow"))?;
            let raw = reader.take(payload_bytes, "dictionary Unicode string")?;
            if reader.read_u16("dictionary Unicode terminator")? != 0 {
                return Err(invalid("Dictionary Unicode string is not terminated"));
            }
            let decoded = decode_utf16(raw, "dictionary Unicode string")?;
            reader.align4(false, "dictionary Unicode padding")?;
            decoded
        } else {
            let raw = reader.take(length - 1, "dictionary string")?;
            if reader.read_u8("dictionary terminator")? != 0 {
                return Err(invalid("Dictionary string is not terminated"));
            }
            decode_ansi(raw, codepage, "dictionary string")?
        };
        if dictionary.insert(identifier, name).is_some() {
            return Err(invalid(format!(
                "Duplicate dictionary property identifier {identifier}"
            )));
        }
        order.push(identifier);
    }
    reader.finish_zero_padding("dictionary padding")?;
    Ok((dictionary, order))
}

fn parse_typed_property(
    data: &[u8],
    codepage: u16,
    property_offset: usize,
) -> Result<PropertyValue, OleError> {
    if data.len() < 4 {
        return Err(invalid("Typed property value is truncated"));
    }
    let variant_type = read_u16(data, 0, "property variant type")?;
    if read_u16(data, 2, "property reserved field")? != 0 {
        return Err(invalid("Typed property reserved field must be zero"));
    }
    let mut reader = ValueReader::new(&data[4..], property_offset + 4);
    let value = parse_value_body(&mut reader, variant_type, codepage, 0)?;
    if reader.remaining_len() <= 3 {
        // Some producers retain up to three opaque bytes after a zero-length
        // top-level value. The property offset table still bounds the value.
        reader.take_remaining();
    } else {
        reader
            .finish_zero_padding("property padding")
            .map_err(|error| {
                invalid(format!(
                    "variant 0x{variant_type:04x} has {} trailing bytes: {error}",
                    reader.remaining_len()
                ))
            })?;
    }
    Ok(value)
}

fn parse_value_body(
    reader: &mut ValueReader<'_>,
    variant_type: u16,
    codepage: u16,
    depth: usize,
) -> Result<PropertyValue, OleError> {
    if depth > 8 {
        return Err(invalid("Property value nesting exceeds the safety limit"));
    }
    if variant_type & VT_VECTOR != 0 {
        let base_type = variant_type & !VT_VECTOR;
        if !is_supported_vector_base(base_type) {
            if depth != 0 {
                return Err(invalid(format!(
                    "Unsupported nested vector type 0x{variant_type:04x}"
                )));
            }
            return Ok(PropertyValue::Unknown {
                variant_type,
                data: reader.take_remaining().to_vec(),
            });
        }
        return parse_vector(reader, base_type, codepage, depth + 1);
    }
    match variant_type {
        VT_EMPTY => Ok(PropertyValue::Empty),
        VT_NULL => Ok(PropertyValue::Null),
        VT_I1 => Ok(PropertyValue::I1(reader.read_i8("I1 value")?)),
        VT_UI1 => Ok(PropertyValue::UI1(reader.read_u8("UI1 value")?)),
        VT_I2 => Ok(PropertyValue::I2(reader.read_i16("I2 value")?)),
        VT_UI2 => Ok(PropertyValue::UI2(reader.read_u16("UI2 value")?)),
        VT_I4 => Ok(PropertyValue::I4(reader.read_i32("I4 value")?)),
        VT_UI4 => Ok(PropertyValue::UI4(reader.read_u32("UI4 value")?)),
        VT_I8 => Ok(PropertyValue::I8(reader.read_i64("I8 value")?)),
        VT_UI8 => Ok(PropertyValue::UI8(reader.read_u64("UI8 value")?)),
        VT_INT => Ok(PropertyValue::Int(reader.read_i32("INT value")?)),
        VT_UINT => Ok(PropertyValue::UInt(reader.read_u32("UINT value")?)),
        VT_R4 => Ok(PropertyValue::R4(f32::from_bits(
            reader.read_u32("R4 value")?,
        ))),
        VT_R8 => Ok(PropertyValue::R8(f64::from_bits(
            reader.read_u64("R8 value")?,
        ))),
        VT_CY => Ok(PropertyValue::Currency(reader.read_i64("CY value")?)),
        VT_DATE => Ok(PropertyValue::Date(f64::from_bits(
            reader.read_u64("DATE value")?,
        ))),
        VT_BSTR => Ok(PropertyValue::Bstr(read_codepage_string(
            reader,
            codepage,
            "BSTR value",
            depth == 0,
        )?)),
        VT_ERROR => Ok(PropertyValue::Error(reader.read_u32("ERROR value")?)),
        VT_BOOL => Ok(PropertyValue::Bool(reader.read_i16("BOOL value")? != 0)),
        VT_DECIMAL => Ok(PropertyValue::Decimal(
            reader
                .take(16, "DECIMAL value")?
                .try_into()
                .expect("checked fixed-size slice"),
        )),
        VT_LPSTR => Ok(PropertyValue::Lpstr(read_codepage_string(
            reader,
            codepage,
            "LPSTR value",
            depth == 0,
        )?)),
        VT_LPWSTR => Ok(PropertyValue::Lpwstr(read_unicode_string(
            reader,
            "LPWSTR value",
            depth == 0,
        )?)),
        VT_FILETIME => Ok(PropertyValue::Filetime(reader.read_u64("FILETIME value")?)),
        VT_BLOB | VT_BLOB_OBJECT => {
            let size = usize::try_from(reader.read_u32("blob size")?)
                .map_err(|_| invalid("Blob size is too large"))?;
            let blob = reader.take(size, "blob data")?.to_vec();
            reader.align4(depth == 0, "blob padding")?;
            Ok(PropertyValue::Blob(blob))
        },
        VT_CF => {
            let size = usize::try_from(reader.read_u32("clipboard size")?)
                .map_err(|_| invalid("Clipboard size is too large"))?;
            if size < 4 {
                return Err(invalid("Clipboard size must include its format field"));
            }
            let format = reader.read_i32("clipboard format")?;
            let payload = reader.take(size - 4, "clipboard data")?.to_vec();
            reader.align4(depth == 0, "clipboard padding")?;
            Ok(PropertyValue::Clipboard {
                format,
                data: payload,
            })
        },
        VT_CLSID => Ok(PropertyValue::Clsid(PropertySetGuid::from_bytes(
            reader
                .take(16, "CLSID value")?
                .try_into()
                .expect("checked fixed-size slice"),
        ))),
        _ if depth == 0 => Ok(PropertyValue::Unknown {
            variant_type,
            data: reader.take_remaining().to_vec(),
        }),
        _ => Err(invalid(format!(
            "Unsupported nested property variant 0x{variant_type:04x}"
        ))),
    }
}

fn parse_vector(
    reader: &mut ValueReader<'_>,
    base_type: u16,
    codepage: u16,
    depth: usize,
) -> Result<PropertyValue, OleError> {
    let count = usize::try_from(reader.read_u32("vector element count")?)
        .map_err(|_| invalid("Vector element count is too large"))?;
    if count > MAX_VECTOR_ELEMENTS {
        return Err(invalid(format!(
            "Vector element count {count} exceeds the safety limit"
        )));
    }
    let minimum = minimum_value_size(base_type).unwrap_or(0);
    if minimum != 0
        && count
            .checked_mul(minimum)
            .is_none_or(|size| size > reader.remaining_len())
    {
        return Err(invalid("Vector element count exceeds its property range"));
    }
    if base_type == VT_VARIANT && count > reader.remaining_len() / 4 {
        return Err(invalid("Variant vector count exceeds its property range"));
    }
    let mut values = Vec::with_capacity(count);
    for _ in 0..count {
        let value = if base_type == VT_VARIANT {
            let nested_type = reader.read_u16("variant vector element type")?;
            if reader.read_u16("variant vector reserved field")? != 0 {
                return Err(invalid("Variant vector reserved field must be zero"));
            }
            let value = parse_value_body(reader, nested_type, codepage, depth + 1)?;
            reader.align4(false, "variant vector element padding")?;
            value
        } else {
            parse_value_body(reader, base_type, codepage, depth + 1)?
        };
        values.push(value);
    }
    Ok(PropertyValue::Vector(values))
}

fn is_supported_vector_base(variant_type: u16) -> bool {
    matches!(
        variant_type,
        VT_I1
            | VT_UI1
            | VT_I2
            | VT_UI2
            | VT_I4
            | VT_UI4
            | VT_I8
            | VT_UI8
            | VT_R4
            | VT_R8
            | VT_CY
            | VT_DATE
            | VT_BSTR
            | VT_ERROR
            | VT_BOOL
            | VT_VARIANT
            | VT_LPSTR
            | VT_LPWSTR
            | VT_FILETIME
            | VT_CF
            | VT_CLSID
    )
}

fn minimum_value_size(variant_type: u16) -> Option<usize> {
    match variant_type {
        VT_I1 | VT_UI1 => Some(1),
        VT_I2 | VT_UI2 | VT_BOOL => Some(2),
        VT_I4 | VT_UI4 | VT_INT | VT_UINT | VT_R4 | VT_ERROR => Some(4),
        VT_I8 | VT_UI8 | VT_R8 | VT_CY | VT_DATE | VT_FILETIME => Some(8),
        VT_CLSID => Some(16),
        VT_VARIANT | VT_BSTR | VT_LPSTR | VT_LPWSTR | VT_CF => Some(4),
        _ => None,
    }
}

fn read_codepage_string(
    reader: &mut ValueReader<'_>,
    codepage: u16,
    description: &str,
    top_level: bool,
) -> Result<String, OleError> {
    let size = usize::try_from(reader.read_u32(description)?)
        .map_err(|_| invalid(format!("{description} is too large")))?;
    let raw = reader.take(size, description)?;
    let value = if size == 0 {
        String::new()
    } else if codepage == UNICODE_CODEPAGE {
        if size % 2 != 0 || !raw.ends_with(&[0, 0]) {
            return Err(invalid(format!("{description} is not terminated UTF-16LE")));
        }
        let end = raw
            .chunks_exact(2)
            .position(|pair| pair == [0, 0])
            .map_or(raw.len(), |units| units * 2);
        decode_utf16(&raw[..end], description)?
    } else {
        if raw.last() != Some(&0) {
            return Err(invalid(format!("{description} is not NUL-terminated")));
        }
        let end = raw.iter().position(|byte| *byte == 0).unwrap_or(raw.len());
        decode_ansi(&raw[..end], codepage, description)?
    };
    reader.align4(top_level, &format!("{description} padding"))?;
    Ok(value)
}

fn read_unicode_string(
    reader: &mut ValueReader<'_>,
    description: &str,
    top_level: bool,
) -> Result<String, OleError> {
    let units = usize::try_from(reader.read_u32(description)?)
        .map_err(|_| invalid(format!("{description} is too large")))?;
    let byte_len = units
        .checked_mul(2)
        .ok_or_else(|| invalid(format!("{description} length overflow")))?;
    let raw = reader.take(byte_len, description)?;
    let value = if units == 0 {
        String::new()
    } else {
        if !raw.ends_with(&[0, 0]) {
            return Err(invalid(format!("{description} is not NUL-terminated")));
        }
        let end = raw
            .chunks_exact(2)
            .position(|pair| pair == [0, 0])
            .map_or(raw.len(), |units| units * 2);
        decode_utf16(&raw[..end], description)?
    };
    reader.align4(top_level, &format!("{description} padding"))?;
    Ok(value)
}

fn decode_utf16(data: &[u8], description: &str) -> Result<String, OleError> {
    if !data.len().is_multiple_of(2) {
        return Err(invalid(format!("{description} has an odd byte length")));
    }
    let units: Vec<u16> = data
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();
    String::from_utf16(&units)
        .map_err(|_| invalid(format!("{description} contains invalid UTF-16")))
}

fn decode_ansi(data: &[u8], codepage: u16, description: &str) -> Result<String, OleError> {
    let page = Mbcs::require(u32::from(codepage))
        .map_err(|error| invalid(format!("Could not decode {description}: {error}")))?;
    page.decode(data)
        .map(Cow::into_owned)
        .map_err(|error| invalid(format!("Could not decode {description}: {error}")))
}

struct ValueReader<'a> {
    data: &'a [u8],
    position: usize,
    alignment_base: usize,
}

impl<'a> ValueReader<'a> {
    const fn new(data: &'a [u8], alignment_base: usize) -> Self {
        Self {
            data,
            position: 0,
            alignment_base,
        }
    }

    fn remaining_len(&self) -> usize {
        self.data.len() - self.position
    }

    fn take(&mut self, length: usize, description: &str) -> Result<&'a [u8], OleError> {
        let start = self.position;
        let end = start
            .checked_add(length)
            .filter(|end| *end <= self.data.len())
            .ok_or_else(|| invalid(format!("{description} exceeds its property range")))?;
        self.position = end;
        Ok(&self.data[start..end])
    }

    fn take_remaining(&mut self) -> &'a [u8] {
        let remaining = &self.data[self.position..];
        self.position = self.data.len();
        remaining
    }

    fn read_u8(&mut self, description: &str) -> Result<u8, OleError> {
        Ok(self.take(1, description)?[0])
    }

    fn read_i8(&mut self, description: &str) -> Result<i8, OleError> {
        Ok(self.read_u8(description)? as i8)
    }

    fn read_u16(&mut self, description: &str) -> Result<u16, OleError> {
        let bytes = self.take(2, description)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn read_i16(&mut self, description: &str) -> Result<i16, OleError> {
        let bytes = self.take(2, description)?;
        Ok(i16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn read_u32(&mut self, description: &str) -> Result<u32, OleError> {
        let bytes = self.take(4, description)?;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("checked fixed-size slice"),
        ))
    }

    fn read_i32(&mut self, description: &str) -> Result<i32, OleError> {
        let bytes = self.take(4, description)?;
        Ok(i32::from_le_bytes(
            bytes.try_into().expect("checked fixed-size slice"),
        ))
    }

    fn read_u64(&mut self, description: &str) -> Result<u64, OleError> {
        let bytes = self.take(8, description)?;
        Ok(u64::from_le_bytes(
            bytes.try_into().expect("checked fixed-size slice"),
        ))
    }

    fn read_i64(&mut self, description: &str) -> Result<i64, OleError> {
        let bytes = self.take(8, description)?;
        Ok(i64::from_le_bytes(
            bytes.try_into().expect("checked fixed-size slice"),
        ))
    }

    fn align4(&mut self, top_level: bool, description: &str) -> Result<(), OleError> {
        let absolute_position = self
            .alignment_base
            .checked_add(self.position)
            .ok_or_else(|| invalid(format!("{description} position overflow")))?;
        let padding = (4 - (absolute_position & 3)) & 3;
        let available = padding.min(self.remaining_len());
        let candidate = &self.data[self.position..self.position + available];
        let consumed = if top_level {
            // Top-level property offsets are authoritative. Several Office
            // producers omit filler or write nonzero filler between values.
            available
        } else {
            // Inside a vector there is no offset table. Match Office readers:
            // skip zero filler only and stop before the next nonzero field.
            candidate.iter().take_while(|byte| **byte == 0).count()
        };
        self.take(consumed, description)?;
        Ok(())
    }

    fn finish_zero_padding(&mut self, description: &str) -> Result<(), OleError> {
        let remaining = &self.data[self.position..];
        if remaining.iter().any(|byte| *byte != 0) {
            return Err(invalid(format!("{description} must be zero")));
        }
        self.position = self.data.len();
        Ok(())
    }
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
        .ok_or_else(|| invalid(format!("{description} exceeds its enclosing range")))?;
    Ok(&data[offset..end])
}

fn read_u16(data: &[u8], offset: usize, description: &str) -> Result<u16, OleError> {
    let bytes = checked_range(data, offset, 2, description)?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, description: &str) -> Result<u32, OleError> {
    let bytes = checked_range(data, offset, 4, description)?;
    Ok(u32::from_le_bytes(
        bytes.try_into().expect("checked fixed-size slice"),
    ))
}

fn read_guid(data: &[u8], offset: usize, description: &str) -> Result<PropertySetGuid, OleError> {
    Ok(PropertySetGuid::from_bytes(
        checked_range(data, offset, 16, description)?
            .try_into()
            .expect("checked fixed-size slice"),
    ))
}

fn invalid(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(message.into())
}

fn filetime_to_date(filetime: u64) -> Option<DateTime<Utc>> {
    const EPOCH_DIFF: i64 = 116_444_736_000_000_000;
    let doc_epoch = i64::try_from(filetime).ok()?;
    let nanos = doc_epoch.checked_sub(EPOCH_DIFF)?.checked_mul(100)?;
    Some(DateTime::from_timestamp_nanos(nanos))
}

fn filetime_to_duration(filetime: u64) -> Option<Duration> {
    let nanos = filetime.checked_mul(100)?;
    Some(Duration::nanoseconds(i64::try_from(nanos).ok()?))
}

fn extract_summary_info(metadata: &mut OleMetadata, section: &PropertySet) {
    if let Some(codepage) = section.codepage {
        metadata.codepage = Some(u32::from(codepage.id()));
    }
    metadata.title = section.property(2).and_then(extract_string);
    metadata.subject = section.property(3).and_then(extract_string);
    metadata.author = section.property(4).and_then(extract_string);
    metadata.keywords = section.property(5).and_then(extract_string);
    metadata.comments = section.property(6).and_then(extract_string);
    metadata.template = section.property(7).and_then(extract_string);
    metadata.last_saved_by = section.property(8).and_then(extract_string);
    metadata.revision_number = section.property(9).and_then(extract_string);
    if let Some(PropertyValue::Filetime(value)) = section.property(10) {
        metadata.edit_time = filetime_to_duration(*value);
    }
    if let Some(PropertyValue::Filetime(value)) = section.property(11) {
        metadata.last_printed_time = filetime_to_date(*value);
    }
    if let Some(PropertyValue::Filetime(value)) = section.property(12) {
        metadata.create_time = filetime_to_date(*value);
    }
    if let Some(PropertyValue::Filetime(value)) = section.property(13) {
        metadata.last_saved_time = filetime_to_date(*value);
    }
    metadata.num_pages = section.property(14).and_then(nonnegative_i4);
    metadata.num_words = section.property(15).and_then(nonnegative_i4);
    metadata.num_chars = section.property(16).and_then(nonnegative_i4);
    metadata.creating_application = section.property(18).and_then(extract_string);
    if let Some(PropertyValue::I4(value)) = section.property(19) {
        metadata.security = Some(*value as u32);
    }
}

fn extract_document_summary_info(metadata: &mut OleMetadata, section: &PropertySet) {
    if metadata.codepage.is_none() {
        metadata.codepage = section.codepage.map(|page| u32::from(page.id()));
    }
    metadata.category = section.property(2).and_then(extract_string);
    metadata.manager = section.property(14).and_then(extract_string);
    metadata.company = section.property(15).and_then(extract_string);
}

fn extract_string(value: &PropertyValue) -> Option<String> {
    match value {
        PropertyValue::Bstr(value) | PropertyValue::Lpstr(value) | PropertyValue::Lpwstr(value) => {
            Some(value.clone())
        },
        _ => None,
    }
}

fn nonnegative_i4(value: &PropertyValue) -> Option<u32> {
    match value {
        PropertyValue::I4(value) => u32::try_from(*value).ok(),
        _ => None,
    }
}

impl From<OleMetadata> for litchi_core::Metadata {
    fn from(ole_metadata: OleMetadata) -> Self {
        litchi_core::Metadata {
            title: ole_metadata.title,
            subject: ole_metadata.subject,
            author: ole_metadata.author,
            keywords: ole_metadata.keywords,
            description: ole_metadata.comments,
            identifier: None,
            language: None,
            template: ole_metadata.template,
            last_modified_by: ole_metadata.last_saved_by,
            revision: ole_metadata.revision_number,
            created: ole_metadata.create_time,
            created_local: None,
            modified: ole_metadata.last_saved_time,
            modified_local: None,
            page_count: ole_metadata.num_pages,
            word_count: ole_metadata.num_words,
            character_count: ole_metadata.num_chars,
            character_count_with_spaces: None,
            editing_time_minutes: None,
            application: ole_metadata.creating_application,
            category: ole_metadata.category,
            company: ole_metadata.company,
            manager: ole_metadata.manager,
            content_status: None,
            content_type: None,
            version: None,
            last_printed_time: ole_metadata.last_printed_time,
            last_printed_local: None,
            last_backup_local: None,
            hyperlink_base: None,
            security: ole_metadata.security,
            codepage: ole_metadata.codepage,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;
    use std::path::Path;

    fn summary_property_stream() -> Vec<u8> {
        let mut data = vec![0u8; 96];
        data[0..2].copy_from_slice(&0xfffeu16.to_le_bytes());
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

    fn fixture(path: &str) -> Vec<u8> {
        std::fs::read(
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../")
                .join(path),
        )
        .unwrap()
    }

    #[test]
    fn parses_typed_stream_and_unsigned_codepage() {
        let stream = PropertySetStream::parse(&summary_property_stream()).unwrap();
        let section = &stream.sections[0];
        assert_eq!(section.page().map(CodePage::id), Some(65001));
        assert_eq!(
            section.property(2),
            Some(&PropertyValue::Lpstr("Hello".into()))
        );
    }

    #[test]
    fn property_set_pages_are_checked_before_mutation() {
        let mut section = PropertySet::new(SUMMARY_INFORMATION_FMTID);
        assert!(section.set_page_id(1201).is_err());
        assert_eq!(section.page(), None);
        assert_eq!(section.property(PID_CODEPAGE), None);
        assert!(section.add(PID_CODEPAGE, PropertyValue::I2(1201)).is_err());

        section.set_page(CodePage::WINDOWS_1252);
        assert_eq!(section.page(), Some(CodePage::WINDOWS_1252));
        assert_eq!(
            section.property(PID_CODEPAGE),
            Some(&PropertyValue::I2(1252))
        );
        assert_eq!(section.clear_page(), Some(CodePage::WINDOWS_1252));
        assert_eq!(section.property(PID_CODEPAGE), None);
    }

    #[test]
    fn rejects_duplicate_offsets_and_truncated_values() {
        let mut duplicate = summary_property_stream();
        duplicate[68..72].copy_from_slice(&24u32.to_le_bytes());
        assert!(PropertySetStream::parse(&duplicate).is_err());

        let mut truncated = summary_property_stream();
        truncated[84..88].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(PropertySetStream::parse(&truncated).is_err());
    }

    #[test]
    fn parses_variant_vectors_and_preserves_unknown_values() {
        let mut vector = Vec::new();
        vector.extend_from_slice(&(VT_VECTOR | VT_VARIANT).to_le_bytes());
        vector.extend_from_slice(&0u16.to_le_bytes());
        vector.extend_from_slice(&2u32.to_le_bytes());
        vector.extend_from_slice(&VT_I4.to_le_bytes());
        vector.extend_from_slice(&0u16.to_le_bytes());
        vector.extend_from_slice(&42i32.to_le_bytes());
        vector.extend_from_slice(&VT_BOOL.to_le_bytes());
        vector.extend_from_slice(&0u16.to_le_bytes());
        vector.extend_from_slice(&(-1i16).to_le_bytes());
        vector.extend_from_slice(&0u16.to_le_bytes());
        assert_eq!(
            parse_typed_property(&vector, DEFAULT_CODEPAGE, 0).unwrap(),
            PropertyValue::Vector(vec![PropertyValue::I4(42), PropertyValue::Bool(true)])
        );

        let unknown = [0x34, 0x12, 0, 0, 1, 2, 3, 4];
        assert_eq!(
            parse_typed_property(&unknown, DEFAULT_CODEPAGE, 0).unwrap(),
            PropertyValue::Unknown {
                variant_type: 0x1234,
                data: vec![1, 2, 3, 4]
            }
        );
    }

    #[test]
    fn reads_apache_poi_named_custom_properties() {
        let bytes = fixture("test-data/poi/test-data/hpsf/TestMickey.doc");
        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let metadata = ole.get_metadata().unwrap();
        assert_eq!(metadata.title.as_deref(), Some("sample title"));
        assert_eq!(metadata.subject.as_deref(), Some("sample subject"));
        assert_eq!(metadata.author.as_deref(), Some("Miroslav Obradovic"));
        assert_eq!(metadata.manager.as_deref(), Some("sample manager"));
        assert_eq!(metadata.company.as_deref(), Some("sample company"));
        assert_eq!(
            metadata.custom_properties.get("Client"),
            Some(&PropertyValue::Lpstr("sample client".into()))
        );
        assert_eq!(
            metadata.custom_properties.get("Division"),
            Some(&PropertyValue::Lpstr("sample division".into()))
        );
    }

    #[test]
    fn reads_apache_poi_two_section_unicode_properties() {
        let bytes = fixture("test-data/poi/test-data/hpsf/TestUnicode.xls");
        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let stream = ole
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
            .unwrap();
        assert_eq!(stream.sections.len(), 2);
        let custom = &stream.sections[1];
        assert_eq!(custom.page(), Some(CodePage::Utf16Le));
        assert_eq!(custom.property(2), Some(&PropertyValue::I4(-96_070_278)));
        assert_eq!(
            custom.property(3),
            Some(&PropertyValue::Lpwstr(
                "MCon_Info zu Office bei Schreiner".into()
            ))
        );
        assert_eq!(
            custom.property(4),
            Some(&PropertyValue::Lpwstr(
                "petrovitsch@schreiner-online.de".into()
            ))
        );
        assert_eq!(
            custom.property(5),
            Some(&PropertyValue::Lpwstr("Petrovitsch, Wilhelm".into()))
        );
    }

    #[test]
    fn projects_existing_document_properties_fixture() {
        let bytes = fixture("test-data/ole/doc/documentProperties.doc");
        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let metadata = ole.get_metadata().unwrap();
        assert_eq!(metadata.title.as_deref(), Some("This is document title"));
        assert_eq!(
            metadata.subject.as_deref(),
            Some("This is document subject")
        );
        assert_eq!(metadata.author.as_deref(), Some("Sergey Vladimirov"));
        assert_eq!(metadata.revision_number.as_deref(), Some("0"));
        assert_eq!(
            metadata.create_time.map(|value| value.timestamp()),
            Some(1_309_939_357)
        );
    }

    #[test]
    fn reads_non_dword_and_zero_length_property_fixtures() {
        let bytes = fixture("test-data/poi/test-data/hpsf/TestNon4ByteBoundary.doc");
        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let metadata = ole.get_metadata().unwrap();
        assert_eq!(
            metadata.creating_application.as_deref(),
            Some("Microsoft Word 10.0")
        );
        assert_eq!(metadata.title.as_deref(), Some(""));
        assert_eq!(metadata.author.as_deref(), Some(""));
        assert_eq!(metadata.company.as_deref(), Some("Cour de Justice"));

        let bytes = fixture("test-data/poi/test-data/hpsf/TestZeroLengthCodePage.mpp");
        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let metadata = ole.get_metadata().unwrap();
        assert_eq!(metadata.creating_application.as_deref(), Some("MSProject"));
        assert_eq!(metadata.title.as_deref(), Some("project1"));
        assert_eq!(metadata.author.as_deref(), Some("Jon Iles"));
        assert_eq!(metadata.company.as_deref(), Some(""));
        let stream = ole
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
            .unwrap();
        assert_eq!(stream.sections.len(), 2);
    }

    #[test]
    fn reads_undefined_filetime_and_word90_fixtures() {
        let bytes = fixture("test-data/poi/test-data/hpsf/TestBug52117.doc");
        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let metadata = ole.get_metadata().unwrap();
        assert_eq!(metadata.last_printed_time, None);
        assert_eq!(
            metadata.edit_time.map(|value| value.num_milliseconds()),
            Some(180_000)
        );

        let bytes = fixture("test-data/poi/test-data/hpsf/TestGermanWord90.doc");
        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let stream = ole
            .property_set_stream(&["\u{0005}SummaryInformation"])
            .unwrap();
        assert_eq!(stream.sections[0].properties.len(), 17);
        let metadata = ole.get_metadata().unwrap();
        assert_eq!(metadata.title.as_deref(), Some("Titel"));
        assert_eq!(metadata.author.as_deref(), Some("Rainer Klute (Autor)"));
        assert_eq!(metadata.subject.as_deref(), Some("Thema"));
    }

    #[test]
    fn rejects_unicode_and_vector_allocation_overflows() {
        let mut unicode = Vec::new();
        unicode.extend_from_slice(&VT_LPWSTR.to_le_bytes());
        unicode.extend_from_slice(&0u16.to_le_bytes());
        unicode.extend_from_slice(&0x4000_0001u32.to_le_bytes());
        assert!(parse_typed_property(&unicode, DEFAULT_CODEPAGE, 0).is_err());

        let mut vector = Vec::new();
        vector.extend_from_slice(&(VT_VECTOR | VT_UI1).to_le_bytes());
        vector.extend_from_slice(&0u16.to_le_bytes());
        vector.extend_from_slice(&u32::MAX.to_le_bytes());
        assert!(parse_typed_property(&vector, DEFAULT_CODEPAGE, 0).is_err());
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
}
