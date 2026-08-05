//! Typed values and bounded in-memory custom-property collections.

use caseless::Caseless;
use chrono::{DateTime, Utc};
use std::collections::BTreeMap;
use unicode_normalization::UnicodeNormalization;

use super::schema::{
    FORMAT_ID, MAX_NAME_BYTES, MAX_NAME_CHARS, MAX_PROPERTIES, MAX_TEXT_BYTES,
    MAX_TOTAL_NAME_BYTES, MAX_TOTAL_TEXT_BYTES, checked_increment, checked_total, invalid, limit,
};
use crate::Result;

/// A custom document-property value.
///
/// `I64` and `F32` preserve vocabulary accepted by older Litchi releases even
/// though Microsoft Office's standard custom-property producer profile uses
/// `I32` and `F64`. Values are never silently narrowed or widened.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    /// An explicit `vt:empty` value.
    Empty,
    /// XML text. New values are written as `vt:lpwstr`; a parsed `vt:lpstr`
    /// retains that wire kind on subsequent writes.
    Text(String),
    /// A signed 32-bit integer (`vt:i4`).
    I32(i32),
    /// A signed 64-bit integer (`vt:i8`).
    I64(i64),
    /// A finite 32-bit float (`vt:r4`).
    F32(f32),
    /// A finite 64-bit float (`vt:r8`).
    F64(f64),
    /// A Boolean (`vt:bool`).
    Bool(bool),
    /// A UTC instant (`vt:filetime`) serialized as RFC3339/XML date-time text.
    Time(DateTime<Utc>),
}

impl From<()> for Value {
    fn from((): ()) -> Self {
        Self::Empty
    }
}

impl From<String> for Value {
    fn from(value: String) -> Self {
        Self::Text(value)
    }
}

impl From<&str> for Value {
    fn from(value: &str) -> Self {
        Self::Text(value.to_owned())
    }
}

impl From<i32> for Value {
    fn from(value: i32) -> Self {
        Self::I32(value)
    }
}

impl From<i64> for Value {
    fn from(value: i64) -> Self {
        Self::I64(value)
    }
}

impl From<f32> for Value {
    fn from(value: f32) -> Self {
        Self::F32(value)
    }
}

impl From<f64> for Value {
    fn from(value: f64) -> Self {
        Self::F64(value)
    }
}

impl From<bool> for Value {
    fn from(value: bool) -> Self {
        Self::Bool(value)
    }
}

impl From<DateTime<Utc>> for Value {
    fn from(value: DateTime<Utc>) -> Self {
        Self::Time(value)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum WireKind {
    Empty,
    Lpstr,
    Lpwstr,
    I4,
    I8,
    R4,
    R8,
    Bool,
    Filetime,
}

impl WireKind {
    pub(crate) fn qualified_name(self) -> &'static str {
        match self {
            Self::Empty => "vt:empty",
            Self::Lpstr => "vt:lpstr",
            Self::Lpwstr => "vt:lpwstr",
            Self::I4 => "vt:i4",
            Self::I8 => "vt:i8",
            Self::R4 => "vt:r4",
            Self::R8 => "vt:r8",
            Self::Bool => "vt:bool",
            Self::Filetime => "vt:filetime",
        }
    }

    pub(crate) fn from_local_name(name: &[u8]) -> Option<Self> {
        match name {
            b"empty" => Some(Self::Empty),
            b"lpstr" => Some(Self::Lpstr),
            b"lpwstr" => Some(Self::Lpwstr),
            b"i4" => Some(Self::I4),
            b"i8" => Some(Self::I8),
            b"r4" => Some(Self::R4),
            b"r8" => Some(Self::R8),
            b"bool" => Some(Self::Bool),
            b"filetime" => Some(Self::Filetime),
            _ => None,
        }
    }

    pub(crate) fn for_value(value: &Value) -> Self {
        match value {
            Value::Empty => Self::Empty,
            Value::Text(_) => Self::Lpwstr,
            Value::I32(_) => Self::I4,
            Value::I64(_) => Self::I8,
            Value::F32(_) => Self::R4,
            Value::F64(_) => Self::R8,
            Value::Bool(_) => Self::Bool,
            Value::Time(_) => Self::Filetime,
        }
    }

    pub(crate) fn is_text(self) -> bool {
        matches!(self, Self::Lpstr | Self::Lpwstr)
    }
}

#[derive(Debug)]
pub(crate) struct Property {
    pub(crate) pid: i32,
    pub(crate) format_id: String,
    pub(crate) wire: WireKind,
    pub(crate) value: Value,
}

/// A bounded collection of custom document properties.
///
/// Names are unique case-insensitively. Exact-name insertion replaces a value
/// while retaining its PID; insertion with different casing is rejected as
/// ambiguous. Iterators use lexical name order, while XML output uses PID order.
#[derive(Debug)]
pub struct Props {
    pub(crate) properties: BTreeMap<String, Property>,
    pub(crate) folded_names: BTreeMap<String, String>,
    pub(crate) next_pid: Option<i32>,
    pub(crate) name_bytes: usize,
    pub(crate) text_bytes: usize,
}

impl Default for Props {
    fn default() -> Self {
        Self::new()
    }
}

impl Props {
    /// Creates an empty property collection.
    #[must_use]
    pub fn new() -> Self {
        Self {
            properties: BTreeMap::new(),
            folded_names: BTreeMap::new(),
            next_pid: Some(2),
            name_bytes: 0,
            text_bytes: 0,
        }
    }

    /// Inserts or replaces a property.
    ///
    /// New PIDs are allocated monotonically with checked arithmetic. The old
    /// value is moved out when an exact-name property is replaced.
    pub fn insert(
        &mut self,
        name: impl Into<String>,
        value: impl Into<Value>,
    ) -> Result<Option<Value>> {
        let name = name.into();
        let value = value.into();
        validate_name(&name)?;
        validate_value(&value)?;

        if let Some(property) = self.properties.get_mut(&name) {
            let old_text = value_text_bytes(&property.value);
            let new_text = value_text_bytes(&value);
            let base = self
                .text_bytes
                .checked_sub(old_text)
                .ok_or_else(|| invalid("custom-property text accounting is inconsistent"))?;
            let updated = checked_total(
                base,
                new_text,
                MAX_TOTAL_TEXT_BYTES,
                "custom-property text bytes",
            )?;
            let wire = if property.wire.is_text() && matches!(value, Value::Text(_)) {
                property.wire
            } else {
                WireKind::for_value(&value)
            };
            self.text_bytes = updated;
            property.wire = wire;
            return Ok(Some(std::mem::replace(&mut property.value, value)));
        }

        let folded = fold_name(&name);
        if self.folded_names.contains_key(&folded) {
            return Err(invalid(format!(
                "custom property name '{name}' duplicates an existing name case-insensitively"
            )));
        }
        let actual_count = checked_increment(self.properties.len(), "custom properties")?;
        if actual_count > MAX_PROPERTIES {
            return Err(limit("custom properties", MAX_PROPERTIES, actual_count));
        }
        let names = checked_total(
            self.name_bytes,
            name.len(),
            MAX_TOTAL_NAME_BYTES,
            "custom-property name bytes",
        )?;
        let texts = checked_total(
            self.text_bytes,
            value_text_bytes(&value),
            MAX_TOTAL_TEXT_BYTES,
            "custom-property text bytes",
        )?;
        let pid = self
            .next_pid
            .ok_or_else(|| invalid("custom property PID space is exhausted"))?;
        self.next_pid = pid.checked_add(1);
        let property = Property {
            pid,
            format_id: FORMAT_ID.to_owned(),
            wire: WireKind::for_value(&value),
            value,
        };
        self.folded_names.insert(folded, name.clone());
        self.properties.insert(name, property);
        self.name_bytes = names;
        self.text_bytes = texts;
        Ok(None)
    }

    /// Borrows a property by its case-insensitive name.
    #[must_use]
    pub fn get(&self, name: &str) -> Option<&Value> {
        self.folded_names
            .get(&fold_name(name))
            .and_then(|stored| self.properties.get(stored))
            .map(|property| &property.value)
    }

    /// Removes a property and moves out its value.
    pub fn remove(&mut self, name: &str) -> Option<Value> {
        let stored = self.folded_names.remove(&fold_name(name))?;
        let property = self.properties.remove(&stored)?;
        self.name_bytes = self.name_bytes.saturating_sub(stored.len());
        self.text_bytes = self
            .text_bytes
            .saturating_sub(value_text_bytes(&property.value));
        Some(property.value)
    }

    /// Returns property names in lexical order.
    pub fn names(&self) -> impl ExactSizeIterator<Item = &str> {
        self.properties.keys().map(String::as_str)
    }

    /// Returns name/value pairs in lexical name order.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = (&str, &Value)> {
        self.properties
            .iter()
            .map(|(name, property)| (name.as_str(), &property.value))
    }

    /// Removes every property and resets PID allocation to 2.
    pub fn clear(&mut self) {
        self.properties.clear();
        self.folded_names.clear();
        self.next_pid = Some(2);
        self.name_bytes = 0;
        self.text_bytes = 0;
    }

    /// Returns whether a case-insensitive property name is present.
    #[must_use]
    pub fn contains(&self, name: &str) -> bool {
        self.folded_names.contains_key(&fold_name(name))
    }

    /// Returns the number of properties.
    #[must_use]
    pub fn len(&self) -> usize {
        self.properties.len()
    }

    /// Returns whether no properties are present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.properties.is_empty()
    }

    pub(crate) fn insert_parsed(
        &mut self,
        name: String,
        pid: i32,
        format_id: String,
        wire: WireKind,
        value: Value,
    ) -> Result<()> {
        validate_name(&name)?;
        validate_value(&value)?;
        if pid < 2 {
            return Err(invalid(format!(
                "custom property '{name}' has PID {pid}; PIDs must be at least 2"
            )));
        }
        let folded = fold_name(&name);
        if self.folded_names.contains_key(&folded) {
            return Err(invalid(format!(
                "duplicate custom property name '{name}' (names are case-insensitive)"
            )));
        }
        let actual_count = checked_increment(self.properties.len(), "custom properties")?;
        if actual_count > MAX_PROPERTIES {
            return Err(limit("custom properties", MAX_PROPERTIES, actual_count));
        }
        let names = checked_total(
            self.name_bytes,
            name.len(),
            MAX_TOTAL_NAME_BYTES,
            "custom-property name bytes",
        )?;
        let texts = checked_total(
            self.text_bytes,
            value_text_bytes(&value),
            MAX_TOTAL_TEXT_BYTES,
            "custom-property text bytes",
        )?;
        self.folded_names.insert(folded, name.clone());
        self.properties.insert(
            name,
            Property {
                pid,
                format_id,
                wire,
                value,
            },
        );
        self.name_bytes = names;
        self.text_bytes = texts;
        self.next_pid = match self.next_pid {
            Some(next) if pid >= next => pid.checked_add(1),
            next => next,
        };
        Ok(())
    }
}

pub(crate) fn validate_value(value: &Value) -> Result<()> {
    match value {
        Value::Text(text) => {
            if text.len() > MAX_TEXT_BYTES {
                return Err(limit(
                    "custom-property text bytes",
                    MAX_TEXT_BYTES,
                    text.len(),
                ));
            }
            validate_xml_text(text, "custom-property text")
        },
        Value::F32(value) if !value.is_finite() => {
            Err(invalid("F32 custom property must be finite"))
        },
        Value::F64(value) if !value.is_finite() => {
            Err(invalid("F64 custom property must be finite"))
        },
        _ => Ok(()),
    }
}

pub(crate) fn validate_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(invalid("custom property name cannot be empty"));
    }
    if name.len() > MAX_NAME_BYTES {
        return Err(limit(
            "custom-property name bytes",
            MAX_NAME_BYTES,
            name.len(),
        ));
    }
    let chars = name.chars().count();
    if chars > MAX_NAME_CHARS {
        return Err(limit(
            "custom-property name characters",
            MAX_NAME_CHARS,
            chars,
        ));
    }
    validate_xml_text(name, "custom-property name")
}

pub(crate) fn validate_xml_text(value: &str, label: &str) -> Result<()> {
    if let Some(character) = value.chars().find(|character| !is_xml_10_char(*character)) {
        return Err(invalid(format!(
            "{label} contains XML 1.0-forbidden character U+{:04X}",
            u32::from(character)
        )));
    }
    Ok(())
}

fn is_xml_10_char(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(u32::from(character), 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

pub(crate) fn fold_name(name: &str) -> String {
    name.chars().nfd().default_case_fold().nfd().collect()
}

pub(crate) fn value_text_bytes(value: &Value) -> usize {
    match value {
        Value::Text(text) => text.len(),
        _ => 0,
    }
}
