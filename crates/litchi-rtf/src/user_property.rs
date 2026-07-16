//! Ordered, inert RTF user-defined document properties.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_USER_PROPERTIES: usize = 65_536;
pub(crate) const MAX_USER_PROPERTY_NAME_BYTES: usize = 65_536;
pub(crate) const MAX_USER_PROPERTY_VALUE_BYTES: usize = 4 * 1_048_576;
pub(crate) const MAX_USER_PROPERTY_TEXT_BYTES: usize = 16 * 1_048_576;

/// The value types defined for `\proptype` by the RTF specification.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum UserPropertyType {
    Integer,
    Real,
    Boolean,
    Text,
    DateTime,
    Unknown(i32),
}

impl UserPropertyType {
    #[must_use]
    pub fn from_code(code: i32) -> Self {
        match code {
            3 => Self::Integer,
            5 => Self::Real,
            11 => Self::Boolean,
            30 => Self::Text,
            64 => Self::DateTime,
            code => Self::Unknown(code),
        }
    }

    #[must_use]
    pub fn code(self) -> i32 {
        match self {
            Self::Integer => 3,
            Self::Real => 5,
            Self::Boolean => 11,
            Self::Text => 30,
            Self::DateTime => 64,
            Self::Unknown(code) => code,
        }
    }
}

/// A validated custom-property date/time.
///
/// LibreOffice emits the RTF dotted date form (`YYYY. MM. DD.`); ISO date and
/// date-time forms are accepted as well. The enclosing value retains the exact
/// lexical spelling for lossless round-tripping.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct UserPropertyDateTime {
    pub year: u16,
    pub month: u8,
    pub day: u8,
    pub hour: Option<u8>,
    pub minute: Option<u8>,
    pub second: Option<u8>,
}

impl UserPropertyDateTime {
    pub fn parse(value: &str) -> RtfResult<Self> {
        let value = value.trim();
        let (date, time) = if let Some(parts) = value.split_once('T') {
            (parts.0, Some(parts.1))
        } else {
            (value, None)
        };
        let date = date.trim_end_matches('.');
        let separator = if date.contains('.') { '.' } else { '-' };
        let mut date_parts = date.split(separator).map(str::trim);
        let year = date_parts.next().and_then(|part| part.parse().ok());
        let month = date_parts.next().and_then(|part| part.parse().ok());
        let day = date_parts.next().and_then(|part| part.parse().ok());
        if date_parts.next().is_some() {
            return Err(RtfError::MalformedDocument(
                "RTF date user property has an invalid date".to_string(),
            ));
        }
        let (hour, minute, second) = if let Some(time) = time {
            let time = time.strip_suffix('Z').unwrap_or(time);
            let mut parts = time.split(':');
            let values = (
                parts.next().and_then(|part| part.parse().ok()),
                parts.next().and_then(|part| part.parse().ok()),
                parts.next().and_then(|part| part.parse().ok()),
            );
            if parts.next().is_some() || values.0.is_none() || values.1.is_none() || values.2.is_none() {
                return Err(RtfError::MalformedDocument(
                    "RTF date user property has an invalid time".to_string(),
                ));
            }
            values
        } else {
            (None, None, None)
        };
        let parsed = Self {
            year: year.ok_or_else(|| RtfError::MalformedDocument(
                "RTF date user property has an invalid year".to_string(),
            ))?,
            month: month.ok_or_else(|| RtfError::MalformedDocument(
                "RTF date user property has an invalid month".to_string(),
            ))?,
            day: day.ok_or_else(|| RtfError::MalformedDocument(
                "RTF date user property has an invalid day".to_string(),
            ))?,
            hour,
            minute,
            second,
        };
        parsed.validate()?;
        Ok(parsed)
    }

    pub fn validate(&self) -> RtfResult<()> {
        if self.year == 0
            || !(1..=12).contains(&self.month)
            || self.hour.is_some_and(|value| value > 23)
            || self.minute.is_some_and(|value| value > 59)
            || self.second.is_some_and(|value| value > 59)
            || self.hour.is_some() != self.minute.is_some()
            || self.minute.is_some() != self.second.is_some()
        {
            return Err(RtfError::MalformedDocument(
                "RTF date user property is outside its valid range".to_string(),
            ));
        }
        let leap = self.year % 4 == 0 && (self.year % 100 != 0 || self.year % 400 == 0);
        let max_day = match self.month {
            2 if leap => 29,
            2 => 28,
            4 | 6 | 9 | 11 => 30,
            _ => 31,
        };
        if self.day == 0 || self.day > max_day {
            return Err(RtfError::MalformedDocument(
                "RTF date user property contains an invalid calendar date".to_string(),
            ));
        }
        Ok(())
    }
}

/// A typed user-defined property value with its original RTF lexical form.
///
/// Values and links are inert metadata. Litchi does not evaluate, refresh, or
/// dereference them.
#[derive(Clone, Debug, PartialEq)]
pub enum UserPropertyValue<'a> {
    Integer { value: i32, lexical: Cow<'a, str> },
    Real { value: f64, lexical: Cow<'a, str> },
    Boolean { value: bool, lexical: Cow<'a, str> },
    Text { value: Cow<'a, str> },
    Date { value: UserPropertyDateTime, lexical: Cow<'a, str> },
    Unknown { type_code: i32, lexical: Cow<'a, str> },
}

impl<'a> UserPropertyValue<'a> {
    /// Parse a value using the RTF 1.9.1 `\proptype` code.
    pub fn from_lexical(
        type_code: i32,
        lexical: impl Into<Cow<'a, str>>,
    ) -> RtfResult<Self> {
        let lexical = lexical.into();
        let value = match type_code {
            3 => Self::Integer {
                value: lexical.parse().map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF integer user property has an invalid value".to_string(),
                    )
                })?,
                lexical,
            },
            5 => {
                let value: f64 = lexical.parse().map_err(|_| {
                    RtfError::MalformedDocument(
                        "RTF real user property has an invalid value".to_string(),
                    )
                })?;
                if !value.is_finite() {
                    return Err(RtfError::MalformedDocument(
                        "RTF real user property must be finite".to_string(),
                    ));
                }
                Self::Real { value, lexical }
            },
            11 => Self::Boolean {
                value: match lexical.as_ref() {
                    "0" => false,
                    "1" => true,
                    _ => {
                        return Err(RtfError::MalformedDocument(
                            "RTF boolean user property must be 0 or 1".to_string(),
                        ));
                    },
                },
                lexical,
            },
            30 => Self::Text { value: lexical },
            64 => Self::Date {
                value: UserPropertyDateTime::parse(&lexical)?,
                lexical,
            },
            type_code => Self::Unknown { type_code, lexical },
        };
        value.validate()?;
        Ok(value)
    }

    /// Return the RTF `\proptype` code.
    pub fn type_code(&self) -> i32 {
        self.property_type().code()
    }

    /// Return the typed RTF property kind.
    #[must_use]
    pub fn property_type(&self) -> UserPropertyType {
        match self {
            Self::Integer { .. } => UserPropertyType::Integer,
            Self::Real { .. } => UserPropertyType::Real,
            Self::Boolean { .. } => UserPropertyType::Boolean,
            Self::Text { .. } => UserPropertyType::Text,
            Self::Date { .. } => UserPropertyType::DateTime,
            Self::Unknown { type_code, .. } => UserPropertyType::Unknown(*type_code),
        }
    }

    /// Return the preserved lexical value written in `\staticval`.
    pub fn lexical(&self) -> &str {
        match self {
            Self::Integer { lexical, .. }
            | Self::Real { lexical, .. }
            | Self::Boolean { lexical, .. }
            | Self::Date { lexical, .. }
            | Self::Unknown { lexical, .. } => lexical,
            Self::Text { value } => value,
        }
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.lexical().len() > MAX_USER_PROPERTY_VALUE_BYTES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF user-property value exceeds {MAX_USER_PROPERTY_VALUE_BYTES} bytes"
            )));
        }
        match self {
            Self::Integer { value, lexical } if lexical.parse::<i32>().ok() != Some(*value) => {
                Err(RtfError::MalformedDocument(
                    "RTF integer user-property lexical and typed values disagree".to_string(),
                ))
            },
            Self::Real { value, lexical }
                if !value.is_finite() || lexical.parse::<f64>().ok() != Some(*value) =>
            {
                Err(RtfError::MalformedDocument(
                    "RTF real user-property lexical and typed values disagree".to_string(),
                ))
            },
            Self::Boolean { value, lexical }
                if lexical.as_ref() != if *value { "1" } else { "0" } =>
            {
                Err(RtfError::MalformedDocument(
                    "RTF boolean user-property lexical and typed values disagree".to_string(),
                ))
            },
            Self::Date { value, lexical }
                if value.validate().is_err()
                    || UserPropertyDateTime::parse(lexical).ok() != Some(*value) =>
            {
                Err(RtfError::MalformedDocument(
                    "RTF date user-property lexical and typed values disagree".to_string(),
                ))
            },
            Self::Unknown { type_code, .. } if matches!(type_code, 3 | 5 | 11 | 30 | 64) => {
                Err(RtfError::MalformedDocument(
                    "known RTF user-property type cannot use Unknown".to_string(),
                ))
            },
            _ => Ok(()),
        }
    }

    pub fn into_owned(self) -> UserPropertyValue<'static> {
        match self {
            Self::Integer { value, lexical } => UserPropertyValue::Integer {
                value,
                lexical: Cow::Owned(lexical.into_owned()),
            },
            Self::Real { value, lexical } => UserPropertyValue::Real {
                value,
                lexical: Cow::Owned(lexical.into_owned()),
            },
            Self::Boolean { value, lexical } => UserPropertyValue::Boolean {
                value,
                lexical: Cow::Owned(lexical.into_owned()),
            },
            Self::Text { value } => UserPropertyValue::Text {
                value: Cow::Owned(value.into_owned()),
            },
            Self::Date { value, lexical } => UserPropertyValue::Date {
                value,
                lexical: Cow::Owned(lexical.into_owned()),
            },
            Self::Unknown { type_code, lexical } => UserPropertyValue::Unknown {
                type_code,
                lexical: Cow::Owned(lexical.into_owned()),
            },
        }
    }
}

/// One named entry in the ordered `\userprops` destination.
#[derive(Clone, Debug, PartialEq)]
pub struct UserProperty<'a> {
    pub name: Cow<'a, str>,
    pub value: UserPropertyValue<'a>,
    pub link_value: Option<Cow<'a, str>>,
}

impl<'a> UserProperty<'a> {
    pub fn new(
        name: impl Into<Cow<'a, str>>,
        value: UserPropertyValue<'a>,
        link_value: Option<impl Into<Cow<'a, str>>>,
    ) -> RtfResult<Self> {
        let property = Self {
            name: name.into(),
            value,
            link_value: link_value.map(Into::into),
        };
        property.validate()?;
        Ok(property)
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.name.is_empty() {
            return Err(RtfError::MalformedDocument(
                "RTF user-property name cannot be empty".to_string(),
            ));
        }
        if self.name.len() > MAX_USER_PROPERTY_NAME_BYTES {
            return Err(RtfError::MalformedDocument(format!(
                "RTF user-property name exceeds {MAX_USER_PROPERTY_NAME_BYTES} bytes"
            )));
        }
        if self
            .link_value
            .as_ref()
            .is_some_and(|link| link.len() > MAX_USER_PROPERTY_VALUE_BYTES)
        {
            return Err(RtfError::MalformedDocument(format!(
                "RTF user-property link exceeds {MAX_USER_PROPERTY_VALUE_BYTES} bytes"
            )));
        }
        self.value.validate()
    }

    pub(crate) fn text_bytes(&self) -> Option<usize> {
        self.name
            .len()
            .checked_add(self.value.lexical().len())?
            .checked_add(self.link_value.as_ref().map_or(0, |link| link.len()))
    }

    pub fn into_owned(self) -> UserProperty<'static> {
        UserProperty {
            name: Cow::Owned(self.name.into_owned()),
            value: self.value.into_owned(),
            link_value: self.link_value.map(|link| Cow::Owned(link.into_owned())),
        }
    }
}
