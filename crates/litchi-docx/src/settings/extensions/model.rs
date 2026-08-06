//! Package-independent Word settings extension values.

use std::fmt::{self, Display, Formatter};
use std::str::FromStr;

use crate::Result;

use super::super::support::{invalid, reserve_one};
use super::validation::validate_extensions;

/// Word 2010 settings-extension namespace.
pub const WORD_2010_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
/// Word 2012 settings-extension namespace.
pub const WORD_2012_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2012/wordml";

/// Maximum number of direct settings extensions retained in one snapshot.
pub const MAX_EXTENSIONS: usize = 128;
/// Maximum bytes retained for one opaque direct-child extension.
pub const MAX_OPAQUE_BYTES: usize = 4 * 1024 * 1024;

/// A fixed-width GUID from the Word 2012 `ST_Guid` type.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Guid([u8; 16]);

impl Guid {
    /// Construct a GUID from its 16-byte wire value.
    #[must_use]
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    /// Borrow the 16-byte wire value.
    #[must_use]
    pub const fn as_bytes(&self) -> &[u8; 16] {
        &self.0
    }

    /// Parse the braced Word `ST_Guid` lexical form.
    pub fn parse(value: &str) -> Result<Self> {
        value.parse()
    }
}

impl Display for Guid {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{{{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            self.0[0],
            self.0[1],
            self.0[2],
            self.0[3],
            self.0[4],
            self.0[5],
            self.0[6],
            self.0[7],
            self.0[8],
            self.0[9],
            self.0[10],
            self.0[11],
            self.0[12],
            self.0[13],
            self.0[14],
            self.0[15],
        )
    }
}

impl FromStr for Guid {
    type Err = crate::Error;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        let bytes = value.as_bytes();
        if bytes.len() != 38
            || bytes.first() != Some(&b'{')
            || bytes.last() != Some(&b'}')
            || [9usize, 14, 19, 24]
                .into_iter()
                .any(|position| bytes.get(position) != Some(&b'-'))
        {
            return Err(invalid(
                "Word settings GUID must be a braced 8-4-4-4-12 value",
            ));
        }

        let mut decoded = [0u8; 16];
        let mut source = 1usize;
        for byte in &mut decoded {
            while matches!(source, 9 | 14 | 19 | 24) {
                source += 1;
            }
            let high = uppercase_hex(
                *bytes
                    .get(source)
                    .ok_or_else(|| invalid("Word settings GUID is truncated"))?,
            )
            .ok_or_else(|| invalid("Word settings GUID contains a non-uppercase-hex digit"))?;
            let low = uppercase_hex(
                *bytes
                    .get(source + 1)
                    .ok_or_else(|| invalid("Word settings GUID is truncated"))?,
            )
            .ok_or_else(|| invalid("Word settings GUID contains a non-uppercase-hex digit"))?;
            *byte = (high << 4) | low;
            source += 2;
        }
        Ok(Self(decoded))
    }
}

fn uppercase_hex(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

/// The two distinct `docId` meanings defined by [MS-DOCX].
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DocumentId {
    /// Word 2010 paragraph-ID context (`w14:docId`).
    ParagraphContext(u32),
    /// Word 2012 source-document identity (`w15:docId`). The GUID attribute
    /// is optional in the schema, so a present element may carry no value.
    Source(Option<Guid>),
}

/// The authored state of a Word `CT_OnOff` settings extension.
///
/// `None` means that the `val` attribute was omitted from a present element;
/// it is therefore distinct from an absent element (`Option<OnOff>::None`)
/// and from an explicit `true` or `false` value.  The schema default for a
/// present element without `val` is on.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub struct OnOff {
    authored: Option<bool>,
}

impl OnOff {
    /// Construct an authored `CT_OnOff` value.
    #[must_use]
    pub const fn new(authored: Option<bool>) -> Self {
        Self { authored }
    }

    /// Construct a present element with the schema-default on value.
    #[must_use]
    pub const fn default_on() -> Self {
        Self::new(None)
    }

    /// Construct a present element with an explicit on value.
    #[must_use]
    pub const fn on() -> Self {
        Self::new(Some(true))
    }

    /// Construct a present element with an explicit off value.
    #[must_use]
    pub const fn off() -> Self {
        Self::new(Some(false))
    }

    /// Return the authored `val` state; omission is `None`.
    #[must_use]
    pub const fn authored(self) -> Option<bool> {
        self.authored
    }

    /// Return the effective value of this present element.
    #[must_use]
    pub const fn effective(self) -> bool {
        match self.authored {
            Some(value) => value,
            None => true,
        }
    }
}

impl From<bool> for OnOff {
    fn from(value: bool) -> Self {
        if value { Self::on() } else { Self::off() }
    }
}

impl From<OnOff> for bool {
    fn from(value: OnOff) -> Self {
        value.effective()
    }
}

impl DocumentId {
    /// Construct a checked Word 2010 paragraph-ID context.
    pub fn paragraph_context(value: u32) -> Result<Self> {
        if value == 0 || value >= 0x8000_0000 {
            return Err(invalid(
                "Word settings paragraph context ID must be greater than 0 and less than 0x80000000",
            ));
        }
        Ok(Self::ParagraphContext(value))
    }

    /// Construct a Word 2012 source-document identity.
    #[must_use]
    pub const fn source(value: Option<Guid>) -> Self {
        Self::Source(value)
    }

    /// Return the paragraph-ID context, when this is the Word 2010 form.
    #[must_use]
    pub const fn paragraph_context_value(&self) -> Option<u32> {
        match self {
            Self::ParagraphContext(value) => Some(*value),
            Self::Source(_) => None,
        }
    }

    /// Return the source GUID, when this is the Word 2012 form.
    #[must_use]
    pub const fn source_value(&self) -> Option<Option<&Guid>> {
        match self {
            Self::ParagraphContext(_) => None,
            Self::Source(value) => Some(value.as_ref()),
        }
    }
}

/// A complete, validated XML element retained without semantic interpretation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OpaqueExtension {
    pub(super) xml: Vec<u8>,
}

impl OpaqueExtension {
    /// Return the original bounded element fragment.
    #[inline]
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }
}

/// One direct extension child of the Word settings element.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Extension {
    /// `w15:chartTrackingRefBased`.
    ChartTrackingRefBased(OnOff),
    /// Either `w14:docId` or `w15:docId`.
    DocumentId(DocumentId),
    /// `w14:conflictMode`.
    ConflictMode(OnOff),
    /// `w14:discardImageEditingData`.
    DiscardImageEditingData(OnOff),
    /// `w14:defaultImageDpi` (`ST_DecimalNumber`).
    DefaultImageDpi(i32),
    /// A direct child whose namespace or local name is not modeled.
    Unknown(OpaqueExtension),
}

impl Extension {
    /// Validate the package-independent value constraints.
    pub fn validate(&self) -> Result<()> {
        match self {
            Self::DocumentId(value) => match value {
                DocumentId::ParagraphContext(value) => {
                    DocumentId::paragraph_context(*value).map(|_| ())
                },
                DocumentId::Source(_) => Ok(()),
            },
            Self::Unknown(value) if value.xml.is_empty() => {
                Err(invalid("opaque settings extension cannot be empty"))
            },
            Self::Unknown(value) if value.xml.len() > MAX_OPAQUE_BYTES => Err(invalid(format!(
                "opaque settings extension exceeds {MAX_OPAQUE_BYTES} bytes"
            ))),
            _ => Ok(()),
        }
    }
}

/// Ordered typed and opaque extensions attached directly to `w:settings`.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Extensions {
    pub(super) values: Vec<Extension>,
}

impl Extensions {
    /// Create an empty extension collection.
    #[must_use]
    pub const fn new() -> Self {
        Self { values: Vec::new() }
    }

    /// Whether no settings extensions are present.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    /// Number of ordered extension children.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    /// Iterate typed and opaque children in their source order.
    #[inline]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = &Extension> {
        self.values.iter()
    }

    /// Validate the complete extension snapshot.
    pub fn validate(&self) -> Result<()> {
        validate_extensions(&self.values)
    }

    /// Append one extension while enforcing count and duplicate constraints.
    pub fn push(&mut self, value: Extension) -> Result<&mut Self> {
        value.validate()?;
        validate_extensions_with_addition(&self.values, &value)?;
        reserve_one(&mut self.values, "Word settings extensions")?;
        self.values.push(value);
        Ok(self)
    }

    /// Remove all extension children.
    pub fn clear(&mut self) {
        self.values.clear();
    }

    /// Return the first chart-reference tracking value.
    pub fn chart_tracking_ref_based(&self) -> Option<OnOff> {
        self.values.iter().find_map(|value| match value {
            Extension::ChartTrackingRefBased(value) => Some(*value),
            _ => None,
        })
    }

    /// Set or remove the chart-reference tracking extension.
    pub fn set_chart_tracking_ref_based(&mut self, value: Option<OnOff>) -> Result<&mut Self> {
        self.replace_unique(value.map(Extension::ChartTrackingRefBased), |value| {
            matches!(value, Extension::ChartTrackingRefBased(_))
        })
    }

    /// Return the Word 2010 paragraph-ID context from `w14:docId`.
    pub fn document_id(&self) -> Option<u32> {
        self.values.iter().find_map(|value| match value {
            Extension::DocumentId(DocumentId::ParagraphContext(value)) => Some(*value),
            _ => None,
        })
    }

    /// Set or remove `w14:docId`.
    pub fn set_document_id(&mut self, value: Option<u32>) -> Result<&mut Self> {
        let value = value
            .map(DocumentId::paragraph_context)
            .transpose()?
            .map(Extension::DocumentId);
        self.replace_unique(value, |value| {
            matches!(
                value,
                Extension::DocumentId(DocumentId::ParagraphContext(_))
            )
        })
    }

    /// Return the optional GUID value from `w15:docId`.
    pub fn source_document_id(&self) -> Option<&Guid> {
        self.values.iter().find_map(|value| match value {
            Extension::DocumentId(DocumentId::Source(Some(value))) => Some(value),
            _ => None,
        })
    }

    /// Set or remove a present `w15:docId` element without a `val`.
    pub fn set_source_document_id_without_value(&mut self, present: bool) -> Result<&mut Self> {
        let value = present.then(|| Extension::DocumentId(DocumentId::source(None)));
        self.replace_unique(value, |value| {
            matches!(value, Extension::DocumentId(DocumentId::Source(_)))
        })
    }

    /// Whether a `w15:docId` element is present, including an absent `val`.
    #[must_use]
    pub fn has_source_document_id(&self) -> bool {
        self.values
            .iter()
            .any(|value| matches!(value, Extension::DocumentId(DocumentId::Source(_))))
    }

    /// Set or remove `w15:docId` with a GUID value.
    pub fn set_source_document_id(&mut self, value: Option<Guid>) -> Result<&mut Self> {
        let value = value
            .map(Some)
            .map(DocumentId::Source)
            .map(Extension::DocumentId);
        self.replace_unique(value, |value| {
            matches!(value, Extension::DocumentId(DocumentId::Source(_)))
        })
    }

    /// Return the conflict-resolution save marker.
    pub fn conflict_mode(&self) -> Option<OnOff> {
        self.values.iter().find_map(|value| match value {
            Extension::ConflictMode(value) => Some(*value),
            _ => None,
        })
    }

    /// Set or remove the conflict-resolution save marker.
    pub fn set_conflict_mode(&mut self, value: Option<OnOff>) -> Result<&mut Self> {
        self.replace_unique(value.map(Extension::ConflictMode), |value| {
            matches!(value, Extension::ConflictMode(_))
        })
    }

    /// Return the image-editing-data discard marker.
    pub fn discard_image_editing_data(&self) -> Option<OnOff> {
        self.values.iter().find_map(|value| match value {
            Extension::DiscardImageEditingData(value) => Some(*value),
            _ => None,
        })
    }

    /// Set or remove the image-editing-data discard marker.
    pub fn set_discard_image_editing_data(&mut self, value: Option<OnOff>) -> Result<&mut Self> {
        self.replace_unique(value.map(Extension::DiscardImageEditingData), |value| {
            matches!(value, Extension::DiscardImageEditingData(_))
        })
    }

    /// Return the default image DPI setting.
    pub fn default_image_dpi(&self) -> Option<i32> {
        self.values.iter().find_map(|value| match value {
            Extension::DefaultImageDpi(value) => Some(*value),
            _ => None,
        })
    }

    /// Set or remove the default image DPI setting.
    pub fn set_default_image_dpi(&mut self, value: Option<i32>) -> Result<&mut Self> {
        self.replace_unique(value.map(Extension::DefaultImageDpi), |value| {
            matches!(value, Extension::DefaultImageDpi(_))
        })
    }

    /// Append a validated opaque direct-child extension.
    pub fn push_unknown(&mut self, value: OpaqueExtension) -> Result<&mut Self> {
        self.push(Extension::Unknown(value))
    }

    /// Return all opaque direct-child extensions in source order.
    pub fn unknown(&self) -> impl Iterator<Item = &OpaqueExtension> {
        self.values.iter().filter_map(|value| match value {
            Extension::Unknown(value) => Some(value),
            _ => None,
        })
    }

    fn replace_unique(
        &mut self,
        value: Option<Extension>,
        matches: impl Fn(&Extension) -> bool,
    ) -> Result<&mut Self> {
        if let Some(index) = self.values.iter().position(|value| matches(value)) {
            if let Some(value) = value {
                value.validate()?;
                self.values[index] = value;
            } else {
                self.values.remove(index);
            }
            return Ok(self);
        }
        if let Some(value) = value {
            self.push(value)?;
        }
        Ok(self)
    }
}

fn validate_extensions_with_addition(values: &[Extension], value: &Extension) -> Result<()> {
    if values.len() >= MAX_EXTENSIONS {
        return Err(invalid(format!(
            "Word settings extension count exceeds {MAX_EXTENSIONS}"
        )));
    }
    if values.iter().any(|existing| same_kind(existing, value)) {
        return Err(invalid("duplicate typed Word settings extension"));
    }
    Ok(())
}

fn same_kind(left: &Extension, right: &Extension) -> bool {
    match (left, right) {
        (Extension::ChartTrackingRefBased(_), Extension::ChartTrackingRefBased(_))
        | (Extension::ConflictMode(_), Extension::ConflictMode(_))
        | (Extension::DiscardImageEditingData(_), Extension::DiscardImageEditingData(_))
        | (Extension::DefaultImageDpi(_), Extension::DefaultImageDpi(_)) => true,
        (
            Extension::DocumentId(DocumentId::ParagraphContext(_)),
            Extension::DocumentId(DocumentId::ParagraphContext(_)),
        )
        | (
            Extension::DocumentId(DocumentId::Source(_)),
            Extension::DocumentId(DocumentId::Source(_)),
        ) => true,
        _ => false,
    }
}
