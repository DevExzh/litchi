//! Semantic values for [MS-OLEPS] alternate-stream binding metadata.

use super::super::binding::Binding;
use super::super::model::Guid;
use litchi_cfb::OleError;
use std::fmt;

/// The required filesystem alternate-stream name for the OLEPS control stream.
///
/// This is a filesystem selector, not a CFB directory name or path.  The
/// control stream's bytes are parsed by [`AlternateStreamControl`].
pub const ALTERNATE_STREAM_CONTROL_NAME: &str = "{4c8cc155-6c1e-11d1-8e41-00c04fb9386d}";

const NON_SIMPLE_PREFIX: &[u8; 5] = b"Docf_";
const MAX_NAME_BYTES: usize = NON_SIMPLE_PREFIX.len() + 27;

/// The fixed [MS-OLEPS] 2.24.3 control-stream packet.
///
/// `Reserved2` is retained when a source packet is parsed, even though the
/// specification says an implementation must ignore it.  Fresh values use
/// zero for both reserved words.  `ApplicationState` and `class_identifier`
/// are application-provided opaque values; this type never interprets either
/// value or activates a class.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AlternateStreamControl {
    pub(crate) reserved2: u16,
    pub(crate) application_state: u32,
    pub(crate) class_identifier: Option<Guid>,
}

impl AlternateStreamControl {
    /// Creates a fresh control packet with both reserved words set to zero.
    #[must_use]
    pub const fn new(application_state: u32, class_identifier: Option<Guid>) -> Self {
        Self {
            reserved2: 0,
            application_state,
            class_identifier,
        }
    }

    /// Returns a copy with a new opaque application state, retaining source
    /// reserved bytes captured by [`Self::parse`].
    #[must_use]
    pub const fn with_application_state(mut self, application_state: u32) -> Self {
        self.application_state = application_state;
        self
    }

    /// Returns a copy with a new optional opaque class identifier, retaining
    /// source reserved bytes captured by [`Self::parse`].
    #[must_use]
    pub const fn with_class_identifier(mut self, class_identifier: Option<Guid>) -> Self {
        self.class_identifier = class_identifier;
        self
    }

    /// Parses exactly one 8-byte or 24-byte control packet.
    ///
    /// # Errors
    ///
    /// Returns an error if the packet has another size or its required-zero
    /// `Reserved1` field is nonzero.
    pub fn parse(bytes: &[u8]) -> Result<Self, OleError> {
        super::codec::decode(bytes)
    }

    /// Serializes the packet as its 8-byte or 24-byte wire form.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        super::codec::encode(self)
    }

    /// Returns the source `Reserved2` word.
    ///
    /// A fresh packet returns zero.  Parsed producer-specific values are
    /// retained for source-preserving replay but have no semantic effect.
    #[must_use]
    pub const fn reserved2(self) -> u16 {
        self.reserved2
    }

    /// Returns the opaque application-provided state value.
    #[must_use]
    pub const fn application_state(self) -> u32 {
        self.application_state
    }

    /// Returns the optional opaque application-provided class identifier.
    #[must_use]
    pub const fn class_identifier(self) -> Option<Guid> {
        self.class_identifier
    }

    pub(super) const fn from_wire(
        reserved2: u16,
        application_state: u32,
        class_identifier: Option<Guid>,
    ) -> Self {
        Self {
            reserved2,
            application_state,
            class_identifier,
        }
    }
}

/// A validated `Docf_` alternate-stream filename for a non-simple property
/// set.
///
/// The value is deliberately separate from
/// [`crate::property_set::BindingName`]: it selects a filesystem alternate
/// stream and must never be passed to a CFB directory lookup.  It stores the
/// canonical `Docf_` prefix followed by the standard OLEPS binding name
/// without allocating during ordinary inspection.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct NonSimpleAlternateStreamName {
    binding: Binding,
    bytes: [u8; MAX_NAME_BYTES],
    len: u8,
}

impl NonSimpleAlternateStreamName {
    /// The exact prefix required by [MS-OLEPS] 2.24.5.
    pub const PREFIX: &'static str = "Docf_";

    /// Creates the canonical non-simple alternate-stream name for `binding`.
    #[must_use]
    pub fn new(binding: Binding) -> Self {
        let binding = canonical_binding(binding);
        let binding_name = binding.name();
        let mut bytes = [0u8; MAX_NAME_BYTES];
        bytes[..NON_SIMPLE_PREFIX.len()].copy_from_slice(NON_SIMPLE_PREFIX);
        bytes[NON_SIMPLE_PREFIX.len()..NON_SIMPLE_PREFIX.len() + binding_name.len()]
            .copy_from_slice(binding_name.as_bytes());
        let len = NON_SIMPLE_PREFIX.len() + binding_name.len();
        Self {
            binding,
            bytes,
            #[allow(
                clippy::cast_possible_truncation,
                reason = "the fixed prefix and BindingName are bounded by MAX_NAME_BYTES"
            )]
            len: len as u8,
        }
    }

    /// Parses exactly one case-insensitive `Docf_` name and its standard OLEPS
    /// binding suffix.
    ///
    /// # Errors
    ///
    /// Returns an error if the prefix is absent or the suffix is not a valid
    /// standard binding name.
    pub fn parse(value: &str) -> Result<Self, OleError> {
        let bytes = value.as_bytes();
        if bytes.len() < NON_SIMPLE_PREFIX.len()
            || !bytes[..NON_SIMPLE_PREFIX.len()].eq_ignore_ascii_case(NON_SIMPLE_PREFIX)
        {
            return Err(super::super::model::invalid(
                "alternate-stream name must start with Docf_",
            ));
        }
        let suffix = &value[NON_SIMPLE_PREFIX.len()..];
        let binding = Binding::from_name(suffix)?;
        Ok(Self::new(binding))
    }

    /// Returns the canonical standard binding represented by this name.
    ///
    /// The wire name cannot distinguish `UserDefinedProperties` from
    /// `DocumentSummaryInformation`, so both aliases return the latter.
    #[must_use]
    pub const fn binding(self) -> Binding {
        self.binding
    }

    /// Returns the canonical filesystem alternate-stream filename.
    ///
    /// This is not a CFB path component.
    #[allow(
        clippy::expect_used,
        reason = "the fixed representation contains only ASCII plus the valid 0x05 prefix"
    )]
    #[must_use]
    pub fn as_str(&self) -> &str {
        std::str::from_utf8(&self.bytes[..usize::from(self.len)])
            .expect("NonSimpleAlternateStreamName stores valid UTF-8 bytes")
    }

    /// Returns the canonical alternate-stream filename bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes[..usize::from(self.len)]
    }

    /// Returns the filename byte length.
    #[must_use]
    pub const fn len(self) -> usize {
        self.len as usize
    }

    /// Whether this validated name is empty.  Valid names are never empty.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.len == 0
    }
}

fn canonical_binding(binding: Binding) -> Binding {
    match Binding::from_format_identifier(binding.format_identifier()) {
        Binding::UserDefinedProperties => Binding::DocumentSummaryInformation,
        canonical => canonical,
    }
}

impl AsRef<str> for NonSimpleAlternateStreamName {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Debug for NonSimpleAlternateStreamName {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("NonSimpleAlternateStreamName")
            .field(&self.as_str())
            .finish()
    }
}

impl fmt::Display for NonSimpleAlternateStreamName {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}
