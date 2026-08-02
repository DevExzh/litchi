//! Inert, positional legacy RTF form-field metadata.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_FORM_FIELDS: usize = 65_536;
pub(crate) const MAX_FORM_FIELD_STRING_BYTES: usize = 65_536;
pub(crate) const MAX_FORM_FIELD_LIST_ENTRIES: usize = 25;
pub(crate) const MAX_FORM_FIELD_DATA_BYTES: usize = 1_048_576;
pub(crate) const MAX_FORM_FIELD_TOTAL_BYTES: usize = 16 * 1_048_576;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormFieldType {
    Text,
    CheckBox,
    DropDown,
}

impl FormFieldType {
    pub(crate) fn from_rtf(value: i32) -> RtfResult<Self> {
        match value {
            0 => Ok(Self::Text),
            1 => Ok(Self::CheckBox),
            2 => Ok(Self::DropDown),
            _ => Err(RtfError::MalformedDocument(
                "RTF form-field type must be 0, 1, or 2".to_string(),
            )),
        }
    }

    pub(crate) const fn to_rtf(self) -> i32 {
        match self {
            Self::Text => 0,
            Self::CheckBox => 1,
            Self::DropDown => 2,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormTextType {
    Regular,
    Number,
    Date,
    CurrentDate,
    CurrentTime,
    Calculation,
}

impl FormTextType {
    pub(crate) fn from_rtf(value: i32) -> RtfResult<Self> {
        match value {
            0 => Ok(Self::Regular),
            1 => Ok(Self::Number),
            2 => Ok(Self::Date),
            3 => Ok(Self::CurrentDate),
            4 => Ok(Self::CurrentTime),
            5 => Ok(Self::Calculation),
            _ => Err(RtfError::MalformedDocument(
                "RTF form text type must be in 0..=5".to_string(),
            )),
        }
    }

    pub(crate) const fn to_rtf(self) -> i32 {
        self as i32
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormField<'a> {
    pub field_type: FormFieldType,
    pub text_type: Option<FormTextType>,
    pub name: Option<Cow<'a, str>>,
    /// Maximum text length; zero means unlimited in RTF.
    pub max_length: Option<u16>,
    /// Inert text-format pattern from `ffformat`.
    pub format: Option<Cow<'a, str>>,
    /// Inert default text from `ffdeftext`.
    pub default_text: Option<Cow<'a, str>>,
    pub default_result: Option<i32>,
    pub result: Option<i32>,
    pub half_point_size: Option<i32>,
    pub protected: bool,
    pub calculate_on_exit: bool,
    pub size_automatically: bool,
    pub own_help: bool,
    pub own_status: bool,
    pub help_text: Option<Cow<'a, str>>,
    pub status_text: Option<Cow<'a, str>>,
    pub entry_macro: Option<Cow<'a, str>>,
    pub exit_macro: Option<Cow<'a, str>>,
    pub list_entries: Vec<Cow<'a, str>>,
    pub has_list_box: bool,
    pub data: Cow<'a, [u8]>,
    pub result_text: Cow<'a, str>,
    pub position: usize,
    pub range_end: usize,
}

impl FormField<'_> {
    pub fn checked(&self) -> Option<bool> {
        (self.field_type == FormFieldType::CheckBox)
            .then_some(self.result)
            .flatten()
            .filter(|value| *value != 25)
            .map(|value| value != 0)
    }

    pub fn default_checked(&self) -> Option<bool> {
        (self.field_type == FormFieldType::CheckBox)
            .then_some(self.default_result)
            .flatten()
            .map(|value| value != 0)
    }

    pub fn selected_index(&self) -> Option<usize> {
        if self.field_type != FormFieldType::DropDown {
            return None;
        }
        self.result
            .filter(|value| *value != 25)
            .and_then(|value| usize::try_from(value).ok())
    }

    pub fn selected_entry(&self) -> Option<&str> {
        self.selected_index()
            .and_then(|index| self.list_entries.get(index))
            .map(Cow::as_ref)
    }

    pub(crate) fn text_bytes(&self) -> Option<usize> {
        let strings = [
            self.name.as_deref(),
            self.format.as_deref(),
            self.default_text.as_deref(),
            self.help_text.as_deref(),
            self.status_text.as_deref(),
            self.entry_macro.as_deref(),
            self.exit_macro.as_deref(),
            Some(self.result_text.as_ref()),
        ];
        let mut total = self.data.len();
        for value in strings.into_iter().flatten() {
            total = total.checked_add(value.len())?;
        }
        for entry in &self.list_entries {
            total = total.checked_add(entry.len())?;
        }
        Some(total)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.range_end < self.position {
            return Err(RtfError::MalformedDocument(
                "RTF form-field range is reversed".to_string(),
            ));
        }
        if self.range_end - self.position != self.result_text.len() {
            return Err(RtfError::MalformedDocument(format!(
                "RTF form-field result range length {} does not match its {} text bytes",
                self.range_end - self.position,
                self.result_text.len()
            )));
        }
        if self.data.len() > MAX_FORM_FIELD_DATA_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF datafield exceeds the safety limit".to_string(),
            ));
        }
        for value in [
            self.name.as_deref(),
            self.format.as_deref(),
            self.default_text.as_deref(),
            self.help_text.as_deref(),
            self.status_text.as_deref(),
            self.entry_macro.as_deref(),
            self.exit_macro.as_deref(),
            Some(self.result_text.as_ref()),
        ]
        .into_iter()
        .flatten()
        .chain(self.list_entries.iter().map(Cow::as_ref))
        {
            if value.len() > MAX_FORM_FIELD_STRING_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF form-field string exceeds the safety limit".to_string(),
                ));
            }
        }
        if self.list_entries.len() > MAX_FORM_FIELD_LIST_ENTRIES {
            return Err(RtfError::MalformedDocument(
                "RTF form-field list exceeds 25 entries".to_string(),
            ));
        }
        if self
            .half_point_size
            .is_some_and(|value| !(1..=32767).contains(&value))
        {
            return Err(RtfError::MalformedDocument(
                "RTF form-field half-point size is outside 1..=32767".to_string(),
            ));
        }
        if self.own_help != self.help_text.is_some()
            || self.own_status != self.status_text.is_some()
        {
            return Err(RtfError::MalformedDocument(
                "RTF form-field own-help/status flags conflict with their text destinations"
                    .to_string(),
            ));
        }

        let valid_index = |value: i32| {
            value == 25
                || usize::try_from(value)
                    .ok()
                    .is_some_and(|index| index < self.list_entries.len())
        };
        match self.field_type {
            FormFieldType::Text => {
                if !self.list_entries.is_empty()
                    || self.has_list_box
                    || self.size_automatically
                    || self.default_result.is_some()
                    || self.result.is_some()
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF text form field contains checkbox/list-only controls".to_string(),
                    ));
                }
            },
            FormFieldType::CheckBox => {
                if self
                    .text_type
                    .is_some_and(|value| value != FormTextType::Regular)
                    || self.max_length.is_some()
                    || self.format.is_some()
                    || self.default_text.is_some()
                    || !self.list_entries.is_empty()
                    || self.has_list_box
                    || self
                        .default_result
                        .is_some_and(|value| !(0..=1).contains(&value))
                    || self
                        .result
                        .is_some_and(|value| value != 25 && !(0..=1).contains(&value))
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF checkbox form-field controls conflict or are out of range".to_string(),
                    ));
                }
            },
            FormFieldType::DropDown => {
                if self
                    .text_type
                    .is_some_and(|value| value != FormTextType::Regular)
                    || self.max_length.is_some()
                    || self.format.is_some()
                    || self.default_text.is_some()
                    || self.size_automatically
                    || !self.has_list_box
                    || self.list_entries.is_empty()
                    || self.default_result.is_some_and(|value| !valid_index(value))
                    || self.result.is_some_and(|value| !valid_index(value))
                {
                    return Err(RtfError::MalformedDocument(
                        "RTF dropdown form-field controls conflict or select an invalid entry"
                            .to_string(),
                    ));
                }
            },
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> FormField<'static> {
        FormField {
            field_type: self.field_type,
            text_type: self.text_type,
            name: self.name.map(|value| Cow::Owned(value.into_owned())),
            max_length: self.max_length,
            format: self.format.map(|value| Cow::Owned(value.into_owned())),
            default_text: self
                .default_text
                .map(|value| Cow::Owned(value.into_owned())),
            default_result: self.default_result,
            result: self.result,
            half_point_size: self.half_point_size,
            protected: self.protected,
            calculate_on_exit: self.calculate_on_exit,
            size_automatically: self.size_automatically,
            own_help: self.own_help,
            own_status: self.own_status,
            help_text: self.help_text.map(|value| Cow::Owned(value.into_owned())),
            status_text: self.status_text.map(|value| Cow::Owned(value.into_owned())),
            entry_macro: self.entry_macro.map(|value| Cow::Owned(value.into_owned())),
            exit_macro: self.exit_macro.map(|value| Cow::Owned(value.into_owned())),
            list_entries: self
                .list_entries
                .into_iter()
                .map(|value| Cow::Owned(value.into_owned()))
                .collect(),
            has_list_box: self.has_list_box,
            data: Cow::Owned(self.data.into_owned()),
            result_text: Cow::Owned(self.result_text.into_owned()),
            position: self.position,
            range_end: self.range_end,
        }
    }
}
