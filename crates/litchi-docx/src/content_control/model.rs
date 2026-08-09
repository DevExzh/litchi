#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
use crate::error::{Error, Result};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;
use std::borrow::Cow;
use std::hash::{Hash, Hasher};

/// Word 2020 SDT data-checksum namespace.
pub const STORE_ITEM_CHECKSUM_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash";

/// Word 2024 SDT formatting-lock namespace.
pub const FORMATTING_ALLOWED_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock";

/// Resource policy for one content-control inventory operation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Limits {
    /// Maximum source XML bytes.
    pub max_input_bytes: usize,
    /// Maximum XML events examined after MCE branch selection.
    pub max_events: usize,
    /// Maximum XML nesting depth.
    pub max_depth: usize,
    /// Maximum content-control occurrences.
    pub max_content_controls: usize,
    /// Maximum data-binding occurrences and MCE namespace/directive entries.
    pub max_bindings: usize,
    /// Maximum list-item occurrences across one parsed inventory.
    pub max_list_items: usize,
    /// Maximum list-item occurrences in one content control.
    pub max_list_items_per_control: usize,
    /// Maximum aggregate bytes copied into semantic strings.
    pub max_metadata_bytes: usize,
    /// Maximum transient XML bytes produced by MCE branch selection.
    pub max_mce_output_bytes: usize,
    /// Maximum transient source-plus-marker bytes used for MCE offset tracking.
    pub max_mce_marked_bytes: usize,
    /// Maximum XML bytes produced by authoring or source-splice mutation.
    pub max_output_bytes: usize,
    /// Maximum exact custom-XML bytes processed by one checksum operation.
    pub max_crc_bytes: usize,
}

impl Limits {
    pub(crate) fn validate(&self) -> Result<()> {
        if [
            self.max_input_bytes,
            self.max_events,
            self.max_depth,
            self.max_content_controls,
            self.max_bindings,
            self.max_list_items,
            self.max_list_items_per_control,
            self.max_metadata_bytes,
            self.max_mce_output_bytes,
            self.max_mce_marked_bytes,
            self.max_output_bytes,
            self.max_crc_bytes,
        ]
        .contains(&0)
        {
            return Err(Error::InvalidFormat(
                "content-control limits must be nonzero".to_string(),
            ));
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: 64 * 1024 * 1024,
            max_events: 1_000_000,
            max_depth: 256,
            max_content_controls: 65_536,
            max_bindings: 65_536,
            max_list_items: 65_536,
            max_list_items_per_control: 32_768,
            max_metadata_bytes: 16 * 1024 * 1024,
            max_mce_output_bytes: 128 * 1024 * 1024,
            max_mce_marked_bytes: 256 * 1024 * 1024,
            max_output_bytes: 128 * 1024 * 1024,
            max_crc_bytes: 16 * 1024 * 1024,
        }
    }
}

/// `WordprocessingML` `ST_Lock` value.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum Lock {
    /// Neither the control nor its contents are locked.
    #[default]
    Unlocked,
    /// The control cannot be deleted.
    SdtLocked,
    /// The contents cannot be edited.
    ContentLocked,
    /// The control cannot be deleted and its contents cannot be edited.
    SdtContentLocked,
}

impl Lock {
    /// Return the exact `WordprocessingML` token.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Unlocked => "unlocked",
            Self::SdtLocked => "sdtLocked",
            Self::ContentLocked => "contentLocked",
            Self::SdtContentLocked => "sdtContentLocked",
        }
    }

    /// Whether the control itself is locked against deletion.
    #[must_use]
    pub const fn locks_control(self) -> bool {
        matches!(self, Self::SdtLocked | Self::SdtContentLocked)
    }

    /// Whether the content is locked against editing.
    #[must_use]
    pub const fn locks_content(self) -> bool {
        matches!(self, Self::ContentLocked | Self::SdtContentLocked)
    }
}

/// Semantic Word 2024 `formattingAllowed` state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FormattingAllowed {
    /// Formatting changes remain permitted while content editing is locked.
    Allowed,
    /// Formatting changes are not permitted.
    Disallowed,
}

impl FormattingAllowed {
    /// Construct from an `ST_OnOff` semantic value.
    #[must_use]
    pub const fn from_bool(value: bool) -> Self {
        if value {
            Self::Allowed
        } else {
            Self::Disallowed
        }
    }

    /// Return the semantic boolean value.
    #[must_use]
    pub const fn as_bool(self) -> bool {
        matches!(self, Self::Allowed)
    }
}

/// A valid Word 2020 custom-XML checksum.
///
/// Equality and hashing use the decoded four bytes. Parsed source lexical text is
/// retained only as provenance, so it does not affect semantic identity.
#[derive(Debug, Clone)]
pub struct Checksum {
    bytes: [u8; 4],
    lexical: Option<Box<str>>,
}

impl PartialEq for Checksum {
    fn eq(&self, other: &Self) -> bool {
        self.bytes == other.bytes
    }
}

impl Eq for Checksum {}

impl Hash for Checksum {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.bytes.hash(state);
    }
}

impl Checksum {
    /// Construct from the four bytes stored by Word.
    #[must_use]
    pub const fn from_bytes(bytes: [u8; 4]) -> Self {
        Self {
            bytes,
            lexical: None,
        }
    }

    /// Parse the strict canonical eight-character Base64 lexical form.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn parse(value: &str) -> Result<Self> {
        if value.len() != 8 || !value.is_ascii() || !value.ends_with("==") {
            return Err(invalid_checksum());
        }
        let mut bytes = [0u8; 4];
        let written = BASE64
            .decode_slice(value.as_bytes(), &mut bytes)
            .map_err(|_source_error| invalid_checksum())?;
        if written != bytes.len() || encode(bytes) != value {
            return Err(invalid_checksum());
        }
        Ok(Self {
            bytes,
            lexical: Some(value.into()),
        })
    }

    /// Compute the checksum over exact, uncompressed and unencrypted part bytes.
    ///
    /// This follows the observed Word interoperability profile: CRC-32 with a
    /// zero seed, represented as a little-endian integer before Base64 encoding.
    /// The profile is not fully determined by the OOXML schema vocabulary alone.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn compute(data: &[u8], limits: &Limits) -> Result<Self> {
        limits.validate()?;
        if data.len() > limits.max_crc_bytes {
            return Err(Error::InvalidFormat(
                "content-control checksum input exceeds the CRC budget".to_string(),
            ));
        }
        Ok(Self::from_word_value(litchi_core::mso_crc32::compute(data)))
    }

    /// Construct from the Word interoperability profile's little-endian integer.
    ///
    /// This convention is an observed Word profile, not an interpretation fixed
    /// solely by the OOXML schema vocabulary.
    #[must_use]
    pub const fn from_word_value(value: u32) -> Self {
        Self::from_bytes(value.to_le_bytes())
    }

    /// Borrow the exact four decoded bytes.
    #[must_use]
    pub const fn as_bytes(&self) -> &[u8; 4] {
        &self.bytes
    }

    /// Return the Word interoperability profile's little-endian integer.
    #[must_use]
    pub const fn word_value(&self) -> u32 {
        u32::from_le_bytes(self.bytes)
    }

    /// Return the original canonical source lexical form when parsed from XML.
    #[must_use]
    pub fn original_lexical(&self) -> Option<&str> {
        self.lexical.as_deref()
    }

    /// Return preserved source lexical text, or canonical authoring text.
    ///
    /// Parsed values borrow their exact source lexical form. Values constructed
    /// in memory have no source provenance and therefore return an owned,
    /// canonical Base64 representation instead.
    #[must_use]
    pub fn lexical(&self) -> Cow<'_, str> {
        self.original_lexical()
            .map_or_else(|| Cow::Owned(self.to_base64()), Cow::Borrowed)
    }

    /// Encode the canonical eight-character Base64 form.
    #[must_use]
    pub fn to_base64(&self) -> String {
        encode(self.bytes)
    }

    /// Compare against exact custom-XML part bytes without evaluating XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn verify(&self, data: &[u8], limits: &Limits) -> Result<ChecksumStatus> {
        let actual = Self::compute(data, limits)?;
        Ok(if self.bytes == actual.bytes {
            ChecksumStatus::Matches
        } else {
            ChecksumStatus::Mismatch {
                expected: self.clone(),
                actual,
            }
        })
    }
}

fn encode(bytes: [u8; 4]) -> String {
    BASE64.encode(bytes)
}

fn invalid_checksum() -> Error {
    Error::InvalidFormat(
        "storeItemChecksum must be canonical Base64 encoding exactly four bytes".to_string(),
    )
}

/// Lossless source representation of a checksum attribute.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum ChecksumValue {
    /// A strict canonical checksum.
    Valid(Checksum),
    /// Malformed lexical text retained for inert inspection and exact no-op preservation.
    Malformed(Box<str>),
}

impl ChecksumValue {
    pub(crate) fn from_source(value: String) -> Self {
        Checksum::parse(&value)
            .map_or_else(|_| Self::Malformed(value.into_boxed_str()), Self::Valid)
    }

    /// Return a valid semantic checksum, if the lexical form was valid.
    #[must_use]
    pub const fn checksum(&self) -> Option<&Checksum> {
        match self {
            Self::Valid(value) => Some(value),
            Self::Malformed(_) => None,
        }
    }

    /// Return preserved source lexical text, or canonical authoring text.
    ///
    /// A valid parsed checksum borrows its original source lexical form. A valid
    /// authored checksum has no source provenance, so this returns its owned,
    /// canonical Base64 representation. Malformed values always borrow their
    /// exact source text.
    #[must_use]
    pub fn lexical(&self) -> Cow<'_, str> {
        match self {
            Self::Valid(value) => value.lexical(),
            Self::Malformed(value) => Cow::Borrowed(value),
        }
    }
}

/// Exact owner vocabulary of one content-control data binding.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum BindingFlavor {
    /// ISO/IEC 29500 `w:dataBinding`, the normative authoring form.
    #[default]
    Core,
    /// Observed Word 2012 `w15:dataBinding` formatted-binding extension.
    Word2012,
}

/// Result of inspecting or verifying checksum metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChecksumStatus {
    /// No checksum attribute is present.
    Absent,
    /// A checksum is present but verification was not requested.
    Unchecked(Checksum),
    /// The checksum lexical form is malformed.
    Malformed(Box<str>),
    /// The checksum matches the exact custom-XML part bytes.
    Matches,
    /// The checksum is valid but does not match the exact part bytes.
    Mismatch {
        /// Value declared by the document.
        expected: Checksum,
        /// Value computed from the part bytes.
        actual: Checksum,
    },
}

/// Inert custom-XML binding metadata attached to an SDT.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataBinding {
    flavor: BindingFlavor,
    xpath: String,
    store_item_id: String,
    prefix_mappings: Option<String>,
    checksum: Option<ChecksumValue>,
}

impl DataBinding {
    /// Construct checked lexical binding metadata. `XPath` is never evaluated.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn new(xpath: impl Into<String>, store_item_id: impl Into<String>) -> Result<Self> {
        let value = Self {
            flavor: BindingFlavor::Core,
            xpath: xpath.into(),
            store_item_id: store_item_id.into(),
            prefix_mappings: None,
            checksum: None,
        };
        super::validate_data_binding_values(&value.xpath, &value.store_item_id, None)?;
        Ok(value)
    }

    pub(crate) fn from_parsed(
        flavor: BindingFlavor,
        xpath: String,
        store_item_id: String,
        prefix_mappings: Option<String>,
        checksum: Option<ChecksumValue>,
    ) -> Self {
        Self {
            flavor,
            xpath,
            store_item_id,
            prefix_mappings,
            checksum,
        }
    }

    /// Construct observed Word 2012 formatted-binding metadata for inspection.
    ///
    /// Canonical MS-DOCX authoring intentionally emits the core flavor.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn word_2012(xpath: impl Into<String>, store_item_id: impl Into<String>) -> Result<Self> {
        let mut value = Self::new(xpath, store_item_id)?;
        value.flavor = BindingFlavor::Word2012;
        Ok(value)
    }

    /// Exact element vocabulary that owned this binding.
    #[must_use]
    pub const fn flavor(&self) -> BindingFlavor {
        self.flavor
    }

    /// Add checked inert namespace-prefix declarations.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn with_prefix_mappings(mut self, mappings: impl Into<String>) -> Result<Self> {
        let mappings = mappings.into();
        super::validate_data_binding_values(&self.xpath, &self.store_item_id, Some(&mappings))?;
        self.prefix_mappings = Some(mappings);
        Ok(self)
    }

    /// Add a valid checksum for canonical authoring.
    #[must_use]
    pub fn with_checksum(mut self, checksum: Checksum) -> Self {
        self.checksum = Some(ChecksumValue::Valid(checksum));
        self
    }

    /// Inert `XPath` lexical text. It is never executed by this API.
    #[must_use]
    pub fn xpath(&self) -> &str {
        &self.xpath
    }

    /// Custom XML store item GUID lexical text.
    #[must_use]
    pub fn store_item_id(&self) -> &str {
        &self.store_item_id
    }

    /// Inert namespace prefix declarations.
    #[must_use]
    pub fn prefix_mappings(&self) -> Option<&str> {
        self.prefix_mappings.as_deref()
    }

    /// Valid checksum value, excluding malformed retained source text.
    #[must_use]
    pub fn checksum(&self) -> Option<&Checksum> {
        self.checksum.as_ref().and_then(ChecksumValue::checksum)
    }

    /// Lossless checksum source state.
    #[must_use]
    pub const fn checksum_value(&self) -> Option<&ChecksumValue> {
        self.checksum.as_ref()
    }

    /// Report checksum state without processing payload bytes.
    #[must_use]
    pub fn checksum_status(&self) -> ChecksumStatus {
        match &self.checksum {
            None => ChecksumStatus::Absent,
            Some(ChecksumValue::Valid(value)) => ChecksumStatus::Unchecked(value.clone()),
            Some(ChecksumValue::Malformed(value)) => ChecksumStatus::Malformed(value.clone()),
        }
    }

    /// Verify against exact custom-XML data-part bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn verify_checksum(&self, data: &[u8], limits: &Limits) -> Result<ChecksumStatus> {
        match &self.checksum {
            None => Ok(ChecksumStatus::Absent),
            Some(ChecksumValue::Malformed(value)) => Ok(ChecksumStatus::Malformed(value.clone())),
            Some(ChecksumValue::Valid(value)) => value.verify(data, limits),
        }
    }
}

/// One source-order `w:sdtPr` occurrence, including controls without IDs.
#[derive(Debug, Clone)]
pub struct Occurrence {
    pub(crate) ordinal: usize,
    pub(crate) control: super::ContentControl,
}

impl Occurrence {
    /// Zero-based source-order occurrence number.
    #[must_use]
    pub const fn ordinal(&self) -> usize {
        self.ordinal
    }

    /// Optional source ID. Missing and duplicate IDs remain distinct occurrences.
    #[must_use]
    pub const fn id(&self) -> Option<u32> {
        self.control.id_opt()
    }

    /// Borrow the semantic control.
    #[must_use]
    pub const fn control(&self) -> &super::ContentControl {
        &self.control
    }
}

/// Bounded, source-order semantic inventory of active content controls.
#[derive(Debug, Clone, Default)]
pub struct Inventory {
    pub(crate) occurrences: Vec<Occurrence>,
}

impl Inventory {
    /// Parse with production defaults and content-control-specific MCE capabilities.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        Self::parse_with_limits(xml, &Limits::default())
    }

    /// Parse with caller-provided resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn parse_with_limits(xml: &[u8], limits: &Limits) -> Result<Self> {
        super::parse_inventory(xml, limits)
    }

    /// Source-order occurrences, including missing and duplicate IDs.
    #[must_use]
    pub fn occurrences(&self) -> &[Occurrence] {
        &self.occurrences
    }

    /// Consume the inventory into semantic controls.
    #[must_use]
    pub fn into_controls(self) -> Vec<super::ContentControl> {
        self.occurrences
            .into_iter()
            .map(|occurrence| occurrence.control)
            .collect()
    }
}
