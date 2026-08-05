//! Semantic OLE Property Set values, sections, streams, and metadata projections.

use chrono::{DateTime, Duration, Utc};
use litchi_cfb::OleError;
use litchi_codepage::Mbcs;
use std::collections::{HashMap, HashSet};
use std::hash::{Hash, Hasher};

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> OleError {
    OleError::Allocation { resource, source }
}

pub(crate) const DEFAULT_CODEPAGE: u16 = 1252;
pub(crate) const UNICODE_CODEPAGE: u16 = 1200;
pub(crate) const PID_DICTIONARY: u32 = 0;
pub(crate) const PID_CODEPAGE: u32 = 1;
pub(crate) const PID_LOCALE: u32 = 0x8000_0000;
pub(crate) const PID_BEHAVIOR: u32 = 0x8000_0003;
pub(crate) const MAX_NAMED_PROPERTY_ID: u32 = 0x7fff_ffff;

pub(crate) fn try_vec_with_capacity<T>(
    capacity: usize,
    resource: &'static str,
) -> Result<Vec<T>, OleError> {
    let mut values = Vec::new();
    values
        .try_reserve_exact(capacity)
        .map_err(|source| allocation(resource, source))?;
    Ok(values)
}

pub(crate) fn try_hash_map_with_capacity<K, V>(
    capacity: usize,
    resource: &'static str,
) -> Result<HashMap<K, V>, OleError>
where
    K: Eq + Hash,
{
    let mut values = HashMap::new();
    values
        .try_reserve(capacity)
        .map_err(|source| allocation(resource, source))?;
    Ok(values)
}

pub(crate) fn try_hash_set_with_capacity<T>(
    capacity: usize,
    resource: &'static str,
) -> Result<HashSet<T>, OleError>
where
    T: Eq + Hash,
{
    let mut values = HashSet::new();
    values
        .try_reserve(capacity)
        .map_err(|source| allocation(resource, source))?;
    Ok(values)
}

pub(crate) fn try_clone_vec<T: Clone>(
    source: &[T],
    resource: &'static str,
) -> Result<Vec<T>, OleError> {
    let mut values = try_vec_with_capacity(source.len(), resource)?;
    values.extend(source.iter().cloned());
    Ok(values)
}

pub(crate) fn try_copy_bytes(source: &[u8], resource: &'static str) -> Result<Vec<u8>, OleError> {
    let mut values = try_vec_with_capacity(source.len(), resource)?;
    values.extend_from_slice(source);
    Ok(values)
}

pub(crate) fn try_clone_string(source: &str, resource: &'static str) -> Result<String, OleError> {
    let mut value = String::new();
    value
        .try_reserve_exact(source.len())
        .map_err(|source| allocation(resource, source))?;
    value.push_str(source);
    Ok(value)
}

pub(crate) fn checked_u32(value: usize, description: &str) -> Result<u32, OleError> {
    u32::try_from(value).map_err(|_| invalid(format!("{description} exceeds u32")))
}

pub(crate) fn checked_add(
    value: usize,
    additional: usize,
    description: &str,
) -> Result<usize, OleError> {
    value
        .checked_add(additional)
        .ok_or_else(|| invalid(format!("{description} overflow")))
}

pub(crate) fn checked_mul(
    value: usize,
    multiplier: usize,
    description: &str,
) -> Result<usize, OleError> {
    value
        .checked_mul(multiplier)
        .ok_or_else(|| invalid(format!("{description} overflow")))
}

pub(crate) fn align4_len(value: usize, description: &str) -> Result<usize, OleError> {
    checked_add(value, 3, description).map(|value| value & !3)
}

pub(crate) fn valid_property_identifier(identifier: u32) -> bool {
    matches!(
        identifier,
        PID_DICTIONARY | PID_CODEPAGE | PID_LOCALE | PID_BEHAVIOR
    ) || (2..=MAX_NAMED_PROPERTY_ID).contains(&identifier)
}

pub(crate) fn valid_named_property_identifier(identifier: u32) -> bool {
    (2..=MAX_NAMED_PROPERTY_ID).contains(&identifier)
}

pub(crate) fn validate_property_name(name: &str) -> Result<(), OleError> {
    if name.is_empty() || name.chars().any(|value| value == '\0') {
        Err(invalid("Property names must be nonempty and NUL-free"))
    } else {
        Ok(())
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(message.into())
}

#[derive(Clone, Copy)]
pub(crate) struct AsciiInsensitive<'a>(pub(crate) &'a str);

impl PartialEq for AsciiInsensitive<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.0.eq_ignore_ascii_case(other.0)
    }
}

impl Eq for AsciiInsensitive<'_> {}

impl Hash for AsciiInsensitive<'_> {
    fn hash<H: Hasher>(&self, state: &mut H) {
        state.write_usize(self.0.len());
        for byte in self.0.bytes() {
            state.write_u8(byte.to_ascii_lowercase());
        }
    }
}

/// A serialized OLE GUID. The byte representation is preserved exactly as stored.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Guid([u8; 16]);

impl Guid {
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
pub struct Stream {
    pub version: u16,
    pub system_identifier: u32,
    pub class_identifier: Guid,
    pub sections: Vec<Section>,
}

/// One independently bounded section within a Property Set stream.
#[derive(Debug, Clone, PartialEq)]
pub struct Section {
    pub format_identifier: Guid,
    pub(crate) codepage: Option<CodePage>,
    pub(crate) dictionary: HashMap<u32, String>,
    pub(crate) properties: HashMap<u32, Value>,
    /// Property descriptor order as stored on disk, including PID 0 when present.
    pub(crate) property_order: Vec<u32>,
    /// Dictionary entry order as stored on disk.
    pub(crate) dictionary_order: Vec<u32>,
}

impl Section {
    pub fn new(format_identifier: Guid) -> Self {
        Self {
            format_identifier,
            codepage: None,
            dictionary: HashMap::new(),
            properties: HashMap::new(),
            property_order: Vec::new(),
            dictionary_order: Vec::new(),
        }
    }
    pub fn property(&self, identifier: u32) -> Option<&Value> {
        self.properties.get(&identifier)
    }

    pub fn property_name(&self, identifier: u32) -> Option<&str> {
        self.dictionary.get(&identifier).map(String::as_str)
    }

    pub fn named_properties(&self) -> impl Iterator<Item = (&str, &Value)> {
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

    pub fn find_named(&self, name: &str) -> Option<(u32, &Value)> {
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

    pub fn add(&mut self, identifier: u32, value: Value) -> Result<(), OleError> {
        if !valid_property_identifier(identifier)
            || matches!(identifier, PID_DICTIONARY | PID_CODEPAGE)
            || self.properties.contains_key(&identifier)
        {
            return Err(invalid(format!(
                "Duplicate or reserved property identifier {identifier}"
            )));
        }
        self.properties
            .try_reserve(1)
            .map_err(|source| allocation("property values", source))?;
        self.property_order
            .try_reserve(1)
            .map_err(|source| allocation("property order", source))?;
        self.properties.insert(identifier, value);
        self.property_order.push(identifier);
        Ok(())
    }

    pub fn add_named(
        &mut self,
        identifier: u32,
        name: String,
        value: Value,
    ) -> Result<(), OleError> {
        validate_property_name(&name)?;
        if !valid_named_property_identifier(identifier) {
            return Err(invalid(format!(
                "Named property identifier {identifier} is outside the normal property range"
            )));
        }
        if self
            .dictionary
            .values()
            .any(|existing| existing.eq_ignore_ascii_case(&name))
        {
            return Err(invalid(format!("Duplicate property name '{name}'")));
        }
        if self.properties.contains_key(&identifier) {
            return Err(invalid(format!(
                "Duplicate or reserved property identifier {identifier}"
            )));
        }
        self.properties
            .try_reserve(1)
            .map_err(|source| allocation("property values", source))?;
        self.property_order
            .try_reserve(1)
            .map_err(|source| allocation("property order", source))?;
        self.dictionary
            .try_reserve(1)
            .map_err(|source| allocation("property dictionary", source))?;
        self.dictionary_order
            .try_reserve(1)
            .map_err(|source| allocation("dictionary order", source))?;
        if !self.property_order.contains(&PID_DICTIONARY) {
            self.property_order
                .try_reserve(1)
                .map_err(|source| allocation("property order", source))?;
        }
        self.properties.insert(identifier, value);
        self.property_order.push(identifier);
        self.dictionary.insert(identifier, name);
        self.dictionary_order.push(identifier);
        if !self.property_order.contains(&PID_DICTIONARY) {
            self.property_order.insert(0, PID_DICTIONARY);
        }
        Ok(())
    }

    pub fn update(&mut self, identifier: u32, value: Value) -> Result<Value, OleError> {
        if identifier == PID_CODEPAGE {
            return Err(invalid("Use set_page to update PID 1"));
        }
        let target = self
            .properties
            .get_mut(&identifier)
            .ok_or_else(|| invalid(format!("Property {identifier} does not exist")))?;
        Ok(std::mem::replace(target, value))
    }

    pub fn replace(&mut self, identifier: u32, value: Value) -> Option<Value> {
        if matches!(identifier, PID_DICTIONARY | PID_CODEPAGE) {
            return None;
        }
        if !self.properties.contains_key(&identifier) {
            self.property_order.push(identifier);
        }
        self.properties.insert(identifier, value)
    }

    pub fn remove(&mut self, identifier: u32) -> Option<Value> {
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

    pub fn remove_named(&mut self, name: &str) -> Option<Value> {
        let identifier = self.dictionary_order.iter().copied().find(|identifier| {
            self.dictionary
                .get(identifier)
                .is_some_and(|value| value.eq_ignore_ascii_case(name))
        })?;
        self.remove(identifier)
    }

    pub fn rename(&mut self, identifier: u32, name: String) -> Result<(), OleError> {
        if identifier == PID_CODEPAGE || !valid_named_property_identifier(identifier) {
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
        if !self.dictionary.contains_key(&identifier) {
            self.dictionary
                .try_reserve(1)
                .map_err(|source| allocation("property dictionary", source))?;
        }
        if !self.dictionary_order.contains(&identifier) {
            self.dictionary_order
                .try_reserve(1)
                .map_err(|source| allocation("dictionary order", source))?;
        }
        if !self.property_order.contains(&PID_DICTIONARY) {
            self.property_order
                .try_reserve(1)
                .map_err(|source| allocation("property order", source))?;
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
        let current = try_hash_set_with_capacity(self.property_order.len(), "property order set")?;
        let mut current = current;
        for identifier in self.property_order.iter().copied() {
            current.insert(identifier);
        }
        let mut proposed = try_hash_set_with_capacity(order.len(), "property reorder set")?;
        for identifier in order.iter().copied() {
            proposed.insert(identifier);
        }
        if current != proposed || proposed.len() != order.len() {
            return Err(invalid(
                "Property reorder must contain every identifier exactly once",
            ));
        }
        self.property_order = try_clone_vec(order, "property order")?;
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
            .insert(PID_CODEPAGE, Value::I2(codepage as i16));
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
pub struct Metadata {
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
    pub custom_properties: HashMap<String, Value>,
}

/// A typed OLE property value. Unsupported variants retain their bounded raw bytes.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
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
    Clsid(Guid),
    Vector(Vec<Value>),
    Unknown { variant_type: u16, data: Vec<u8> },
}

pub(crate) fn try_clone_property_value(value: &Value) -> Result<Value, OleError> {
    Ok(match value {
        Value::Empty => Value::Empty,
        Value::Null => Value::Null,
        Value::I1(value) => Value::I1(*value),
        Value::UI1(value) => Value::UI1(*value),
        Value::I2(value) => Value::I2(*value),
        Value::UI2(value) => Value::UI2(*value),
        Value::I4(value) => Value::I4(*value),
        Value::UI4(value) => Value::UI4(*value),
        Value::I8(value) => Value::I8(*value),
        Value::UI8(value) => Value::UI8(*value),
        Value::Int(value) => Value::Int(*value),
        Value::UInt(value) => Value::UInt(*value),
        Value::R4(value) => Value::R4(*value),
        Value::R8(value) => Value::R8(*value),
        Value::Currency(value) => Value::Currency(*value),
        Value::Date(value) => Value::Date(*value),
        Value::Bstr(value) => Value::Bstr(try_clone_string(value, "property value string")?),
        Value::Error(value) => Value::Error(*value),
        Value::Bool(value) => Value::Bool(*value),
        Value::Decimal(value) => Value::Decimal(*value),
        Value::Lpstr(value) => Value::Lpstr(try_clone_string(value, "property value string")?),
        Value::Lpwstr(value) => Value::Lpwstr(try_clone_string(value, "property value string")?),
        Value::Filetime(value) => Value::Filetime(*value),
        Value::Blob(value) => Value::Blob(try_copy_bytes(value, "property value blob")?),
        Value::Clipboard { format, data } => Value::Clipboard {
            format: *format,
            data: try_copy_bytes(data, "property value clipboard data")?,
        },
        Value::Clsid(value) => Value::Clsid(*value),
        Value::Vector(values) => {
            let mut cloned = try_vec_with_capacity(values.len(), "property value vector")?;
            for value in values {
                cloned.push(try_clone_property_value(value)?);
            }
            Value::Vector(cloned)
        },
        Value::Unknown { variant_type, data } => Value::Unknown {
            variant_type: *variant_type,
            data: try_copy_bytes(data, "unknown property value")?,
        },
    })
}

pub(crate) fn try_clone_property_set(section: &Section) -> Result<Section, OleError> {
    let mut dictionary =
        try_hash_map_with_capacity(section.dictionary.len(), "property dictionary")?;
    for (identifier, name) in &section.dictionary {
        dictionary.insert(
            *identifier,
            try_clone_string(name, "property dictionary name")?,
        );
    }

    let mut properties = try_hash_map_with_capacity(section.properties.len(), "property values")?;
    for (identifier, value) in &section.properties {
        properties.insert(*identifier, try_clone_property_value(value)?);
    }

    Ok(Section {
        format_identifier: section.format_identifier,
        codepage: section.codepage,
        dictionary,
        properties,
        property_order: try_clone_vec(&section.property_order, "property order")?,
        dictionary_order: try_clone_vec(&section.dictionary_order, "dictionary order")?,
    })
}

impl Stream {
    /// Version 0 property sets use only the original OLE Property Set types.
    pub const VERSION_0: u16 = 0;
    /// Version 1 property sets add the versioned numeric families.
    pub const VERSION_1: u16 = 1;

    pub fn new(section: Section) -> Self {
        Self {
            version: Self::VERSION_0,
            system_identifier: 0,
            class_identifier: Guid::from_bytes([0; 16]),
            sections: vec![section],
        }
    }
    pub fn section(&self, format_identifier: Guid) -> Option<&Section> {
        self.sections
            .iter()
            .find(|section| section.format_identifier == format_identifier)
    }
    pub fn section_mut(&mut self, format_identifier: Guid) -> Option<&mut Section> {
        self.sections
            .iter_mut()
            .find(|section| section.format_identifier == format_identifier)
    }
    pub fn add_section(&mut self, section: Section) -> Result<(), OleError> {
        if self.sections.len() == 2 || self.section(section.format_identifier).is_some() {
            return Err(invalid("Duplicate or excess Property Set section"));
        }
        self.sections
            .try_reserve(1)
            .map_err(|source| allocation("property-set sections", source))?;
        self.sections.push(section);
        Ok(())
    }
    pub fn remove_section(&mut self, format_identifier: Guid) -> Option<Section> {
        let index = self
            .sections
            .iter()
            .position(|section| section.format_identifier == format_identifier)?;
        Some(self.sections.remove(index))
    }
    pub fn reorder_sections(&mut self, order: &[Guid]) -> Result<(), OleError> {
        if order.len() != self.sections.len() {
            return Err(invalid("Section reorder is incomplete or duplicated"));
        }
        let mut proposed = try_hash_set_with_capacity(order.len(), "section reorder set")?;
        for identifier in order.iter().copied() {
            proposed.insert(identifier);
        }
        if proposed.len() != order.len() {
            return Err(invalid("Section reorder is incomplete or duplicated"));
        }
        let mut reordered = try_vec_with_capacity(order.len(), "property-set sections")?;
        for id in order {
            let index = self
                .sections
                .iter()
                .position(|section| section.format_identifier == *id)
                .ok_or_else(|| invalid("Section reorder references an unknown format ID"))?;
            reordered.push(try_clone_property_set(&self.sections[index])?);
        }
        self.sections = reordered;
        Ok(())
    }
    pub fn clear_sections(&mut self) {
        self.sections.clear();
    }
}

pub const SUMMARY_INFORMATION_FMTID: Guid = Guid::from_bytes([
    0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9,
]);
pub const DOCUMENT_SUMMARY_INFORMATION_FMTID: Guid = Guid::from_bytes([
    0x02, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE,
]);
pub const USER_DEFINED_PROPERTIES_FMTID: Guid = Guid::from_bytes([
    0x05, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE,
]);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Standard {
    SummaryInformation,
    DocumentSummaryInformation,
    UserDefinedProperties,
}

impl Standard {
    pub(crate) fn path(self) -> &'static str {
        match self {
            Self::SummaryInformation => "\u{0005}SummaryInformation",
            Self::DocumentSummaryInformation | Self::UserDefinedProperties => {
                "\u{0005}DocumentSummaryInformation"
            },
        }
    }
    pub(crate) fn format_id(self) -> Guid {
        match self {
            Self::SummaryInformation => SUMMARY_INFORMATION_FMTID,
            Self::DocumentSummaryInformation => DOCUMENT_SUMMARY_INFORMATION_FMTID,
            Self::UserDefinedProperties => USER_DEFINED_PROPERTIES_FMTID,
        }
    }
}

impl From<Metadata> for litchi_core::Metadata {
    fn from(ole_metadata: Metadata) -> Self {
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
