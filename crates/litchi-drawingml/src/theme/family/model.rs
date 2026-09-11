//! Typed values for the DrawingML 2012 `themeFamily` fragment.

use std::{fmt, sync::Arc};

use thiserror::Error as ThisError;

use crate::{Error, Result};

use super::{MAX_ATTRIBUTE_VALUE_BYTES, MAX_GUID_BYTES, MAX_NAME_BYTES};

/// A checked `a:ST_Guid` lexical value.
///
/// The schema's lexical form is a braced, hyphenated hexadecimal GUID.  The
/// The validated uppercase lexical value is retained; XML Schema token edge
/// whitespace is removed while constructing the typed scalar. Untouched
/// producer spelling remains available through the source-backed family.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[must_use]
pub struct Guid(Arc<str>);

impl Guid {
    /// Construct a bounded braced `ST_Guid` value.
    ///
    /// # Errors
    ///
    /// Returns [`ValueError::Guid`] for a value outside the schema lexical
    /// form and [`ValueError::TooLong`] when the explicit bound is exceeded.
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, ValueError> {
        let value = value.as_ref();
        // Bound the borrowed input before any error or normalized-value
        // allocation. XML token edge whitespace may surround the 38-byte
        // lexical value, so the raw bound is deliberately wider than
        // `MAX_GUID_BYTES` while remaining finite.
        if value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(ValueError::TooLong {
                field: "theme family GUID input",
                limit: MAX_ATTRIBUTE_VALUE_BYTES,
            });
        }
        // `ST_Guid` derives from `xsd:token`, whose XML whitespace domain is
        // limited to space, tab, CR, and LF.  Only those edge characters are
        // collapsed here; any inner or non-XML whitespace remains invalid.
        let normalized =
            value.trim_matches(|character| matches!(character, ' ' | '\t' | '\r' | '\n'));
        if normalized.len() > MAX_GUID_BYTES {
            return Err(ValueError::TooLong {
                field: "theme family GUID",
                limit: MAX_GUID_BYTES,
            });
        }
        if !is_guid(normalized) {
            return Err(ValueError::Guid {
                value: value.to_owned(),
            });
        }
        Ok(Self(Arc::from(normalized)))
    }

    /// Borrow the validated GUID lexical value without XML token edge whitespace.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for Guid {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for Guid {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl TryFrom<&str> for Guid {
    type Error = ValueError;

    fn try_from(value: &str) -> std::result::Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl TryFrom<String> for Guid {
    type Error = ValueError;

    fn try_from(value: String) -> std::result::Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Guid> for String {
    fn from(value: Guid) -> Self {
        value.0.as_ref().to_owned()
    }
}

/// Failure to construct a typed Theme Family scalar.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum ValueError {
    /// A GUID is not in the braced `ST_Guid` lexical form.
    #[error("invalid theme family GUID '{value}'")]
    Guid { value: String },
    /// A scalar exceeds its explicit bound.
    #[error("theme family {field} exceeds the limit of {limit} bytes")]
    TooLong { field: &'static str, limit: usize },
}

impl From<ValueError> for Error {
    fn from(error: ValueError) -> Self {
        Self::Invalid(error.to_string())
    }
}

/// Typed projection of `thm15:themeFamily`.
///
/// `name`, `id`, and `vid` are the three required schema attributes from
/// `[MS-ODRAWXML]` §2.4.3.1.  A parsed value retains one immutable source
/// allocation.  The source remains available after typed setters so a later
/// write can patch only these known attributes and retain unknown attributes,
/// namespace declarations, extension children, comments, and lexical forms.
#[derive(Debug, Clone)]
#[must_use]
pub struct Family {
    name: Arc<str>,
    id: Guid,
    variant_id: Guid,
    source: Option<Arc<Source>>,
}

#[derive(Debug, Clone)]
pub(crate) struct Source {
    pub(crate) xml: Arc<[u8]>,
    pub(crate) root_start: usize,
    pub(crate) root_end: usize,
    pub(crate) name: Arc<str>,
    pub(crate) id: Guid,
    pub(crate) variant_id: Guid,
}

impl PartialEq for Family {
    fn eq(&self, other: &Self) -> bool {
        self.name == other.name && self.id == other.id && self.variant_id == other.variant_id
    }
}

impl Eq for Family {}

impl Family {
    /// Construct a detached Theme Family value.
    ///
    /// The input can be either strings or already-checked [`Guid`] values.
    /// `name` follows the XML Schema `xsd:string` domain and may be empty, but
    /// remains explicitly bounded for safe detached authoring.
    ///
    /// # Errors
    ///
    /// Returns an error when a name or GUID exceeds its bound or contains an
    /// XML 1.0 forbidden character.
    pub fn new(
        name: impl AsRef<str>,
        id: impl AsRef<str>,
        variant_id: impl AsRef<str>,
    ) -> Result<Self> {
        let name = checked_name(name.as_ref())?;
        let id = Guid::new(id).map_err(Error::from)?;
        let variant_id = Guid::new(variant_id).map_err(Error::from)?;
        Ok(Self {
            name,
            id,
            variant_id,
            source: None,
        })
    }

    /// Borrow the applied theme name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the applied theme GUID.
    #[must_use]
    pub fn id(&self) -> &Guid {
        &self.id
    }

    /// Borrow the applied variant GUID.
    #[must_use]
    pub fn variant_id(&self) -> &Guid {
        &self.variant_id
    }

    /// Return the retained namespace-complete, standalone-ready source.
    ///
    /// For values parsed by [`super::read`], these are the exact input bytes.
    /// Projections read from a complete Theme part may include injected `xmlns`
    /// declarations for inherited bindings, so their source can differ from
    /// the raw fragment in that part. Use [`super::part::Snapshot::xml_bytes`]
    /// with [`super::part::Snapshot::family_range`] for those exact raw bytes.
    /// Scalar edits retain this source; serialization applies the staged values.
    #[must_use]
    pub fn source(&self) -> Option<&[u8]> {
        self.source.as_deref().map(|source| source.xml.as_ref())
    }

    /// Replace the applied theme name.
    ///
    /// The source allocation is retained; serialization patches the source
    /// attribute and therefore keeps all unmodeled source content.
    pub fn set_name(&mut self, name: impl AsRef<str>) -> Result<&mut Self> {
        self.name = checked_name(name.as_ref())?;
        Ok(self)
    }

    /// Replace the applied theme GUID.
    pub fn set_id(&mut self, id: impl AsRef<str>) -> Result<&mut Self> {
        self.id = Guid::new(id).map_err(Error::from)?;
        Ok(self)
    }

    /// Replace the applied variant GUID.
    pub fn set_variant_id(&mut self, variant_id: impl AsRef<str>) -> Result<&mut Self> {
        self.variant_id = Guid::new(variant_id).map_err(Error::from)?;
        Ok(self)
    }

    /// Builder-style replacement of the applied theme name.
    pub fn with_name(mut self, name: impl AsRef<str>) -> Result<Self> {
        self.set_name(name)?;
        Ok(self)
    }

    /// Builder-style replacement of the applied theme GUID.
    pub fn with_id(mut self, id: impl AsRef<str>) -> Result<Self> {
        self.set_id(id)?;
        Ok(self)
    }

    /// Builder-style replacement of the applied variant GUID.
    pub fn with_variant_id(mut self, variant_id: impl AsRef<str>) -> Result<Self> {
        self.set_variant_id(variant_id)?;
        Ok(self)
    }

    pub(crate) fn from_source(source: Source) -> Self {
        Self {
            name: source.name.clone(),
            id: source.id.clone(),
            variant_id: source.variant_id.clone(),
            source: Some(Arc::new(source)),
        }
    }

    pub(crate) fn source_state(&self) -> Option<&Source> {
        self.source.as_deref()
    }

    pub(crate) fn detached(&self) -> Self {
        Self {
            name: self.name.clone(),
            id: self.id.clone(),
            variant_id: self.variant_id.clone(),
            source: None,
        }
    }

    pub(crate) fn semantic_eq(&self, other: &Self) -> bool {
        self == other
    }
}

pub(crate) fn checked_name(value: &str) -> Result<Arc<str>> {
    if value.len() > MAX_NAME_BYTES {
        return Err(Error::Limit {
            resource: "theme family name bytes",
            limit: MAX_NAME_BYTES,
        });
    }
    validate_xml_text(value, "theme family name")?;
    Ok(Arc::from(value))
}

pub(crate) fn validate_xml_text(value: &str, field: &str) -> Result<()> {
    if let Some(character) = value.chars().find(|character| !is_xml_1_0(*character)) {
        return Err(Error::Invalid(format!(
            "{field} contains an XML 1.0 forbidden character U+{:04X}",
            u32::from(character)
        )));
    }
    Ok(())
}

fn is_xml_1_0(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(character, '\u{20}'..='\u{D7FF}')
        || matches!(character, '\u{E000}'..='\u{FFFD}')
        || matches!(character, '\u{10000}'..='\u{10FFFF}')
}

fn is_guid(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == MAX_GUID_BYTES
        && bytes[0] == b'{'
        && bytes[37] == b'}'
        && bytes[9] == b'-'
        && bytes[14] == b'-'
        && bytes[19] == b'-'
        && bytes[24] == b'-'
        && bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 0 | 9 | 14 | 19 | 24 | 37)
                || byte.is_ascii_digit()
                || matches!(byte, b'A'..=b'F')
        })
}
