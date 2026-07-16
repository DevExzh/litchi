//! Ordered, inert RTF user-defined document properties.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_USER_PROPERTIES: usize = 65_536;
pub(crate) const MAX_USER_PROPERTY_NAME_BYTES: usize = 65_536;
pub(crate) const MAX_USER_PROPERTY_VALUE_BYTES: usize = 4 * 1_048_576;
pub(crate) const MAX_USER_PROPERTY_TEXT_BYTES: usize = 16 * 1_048_576;

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
    Date { lexical: Cow<'a, str> },
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
            64 => {
                if lexical.is_empty() {
                    return Err(RtfError::MalformedDocument(
                        "RTF date user property cannot be empty".to_string(),
                    ));
                }
                Self::Date { lexical }
            },
            type_code => Self::Unknown { type_code, lexical },
        };
        value.validate()?;
        Ok(value)
    }

    /// Return the RTF `\proptype` code.
    pub fn type_code(&self) -> i32 {
        match self {
            Self::Integer { .. } => 3,
            Self::Real { .. } => 5,
            Self::Boolean { .. } => 11,
            Self::Text { .. } => 30,
            Self::Date { .. } => 64,
            Self::Unknown { type_code, .. } => *type_code,
        }
    }

    /// Return the preserved lexical value written in `\staticval`.
    pub fn lexical(&self) -> &str {
        match self {
            Self::Integer { lexical, .. }
            | Self::Real { lexical, .. }
            | Self::Boolean { lexical, .. }
            | Self::Date { lexical }
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
            Self::Date { lexical } if lexical.is_empty() => Err(RtfError::MalformedDocument(
                "RTF date user property cannot be empty".to_string(),
            )),
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
            Self::Date { lexical } => UserPropertyValue::Date {
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
