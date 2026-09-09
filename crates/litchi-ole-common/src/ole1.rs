//! Bounded, inert OLE 1.0 object and presentation streams.
//!
//! This module owns the byte sequence described by [MS-OLEDS] section 2.2.
//! It deliberately does not open linked paths, decode native documents, render
//! presentations, or activate an OLE class.  The borrowed parser retains no
//! owned strings or payloads.  The owned snapshot keeps one immutable source
//! allocation and edits publish a reparsed candidate only after all declared
//! sizes and output limits have been checked.
//!
//! The normative OLE1 header strings are encoded ANSI strings.  Registered
//! clipboard presentation names may be ANSI or UTF-16LE.  Their encoded bytes
//! are exposed and retained; no code page or Unicode normalization is guessed.

use litchi_cfb::OleError;
use std::ops::Range;
use std::sync::Arc;

/// The OLE1 `ObjectHeader` format identifier for a linked object.
pub const FORMAT_ID_LINKED: u32 = 0x0000_0001;
/// The OLE1 `ObjectHeader` format identifier for an embedded object.
pub const FORMAT_ID_EMBEDDED: u32 = 0x0000_0002;
/// The OLE1 presentation format identifier for a header with no class name.
pub const PRESENTATION_FORMAT_ID_NONE: u32 = 0x0000_0000;
/// The OLE1 presentation format identifier for a header with a class name.
pub const PRESENTATION_FORMAT_ID_CLASS: u32 = 0x0000_0005;

/// The standard OLE1 class name for a Windows metafile presentation.
pub const CLASS_METAFILEPICT: &[u8] = b"METAFILEPICT";
/// The standard OLE1 class name for a Bitmap16 presentation.
pub const CLASS_BITMAP: &[u8] = b"BITMAP";
/// The standard OLE1 class name for a device-independent bitmap presentation.
pub const CLASS_DIB: &[u8] = b"DIB";

/// Standard clipboard format `CF_BITMAP`.
pub const CF_BITMAP: u32 = 0x0002;
/// Standard clipboard format `CF_METAFILEPICT`.
pub const CF_METAFILEPICT: u32 = 0x0003;
/// Standard clipboard format `CF_DIB`.
pub const CF_DIB: u32 = 0x0008;
/// Standard clipboard format `CF_ENHMETAFILE`.
pub const CF_ENHMETAFILE: u32 = 0x000e;

const DEFAULT_MAX_BYTES: usize = 64 * 1024 * 1024;
const HARD_MAX_BYTES: usize = 512 * 1024 * 1024;
const DEFAULT_MAX_STRING_BYTES: usize = 1024 * 1024;
const DEFAULT_MAX_REGISTERED_BYTES: usize = 64 * 1024;
const CANONICAL_OLE_VERSION: u32 = 0x0000_0501;

/// A caller-selected compatibility profile for source reads.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Compatibility {
    /// Admit only the structures described by the normative OLE1 grammar.
    #[default]
    Strict,
    /// Admit the narrow producer deviation that appends exactly one
    /// FormatID-zero presentation header after an embedded native payload.
    ///
    /// The eight-byte header is a source boundary marker; the native payload
    /// remains opaque and is not interpreted as RTF (or any other format).
    /// The profile is read/preserve-only and never authors a FormatID-zero
    /// presentation.
    EmptyPresentationHeader,
    /// Admit an embedded object whose native payload reaches the end of the
    /// source without any Presentation structure.
    ///
    /// LibreOffice's ReqIF OLE producer emits this source shape for a native
    /// Draw document.  The profile is read/preserve-only; a typed presentation
    /// edit remains refused and canonical authoring always emits one of the
    /// five normative presentation forms.
    MissingPresentation,
}

impl Compatibility {
    const fn allows_empty_presentation_header(self) -> bool {
        matches!(self, Self::EmptyPresentationHeader)
    }

    const fn allows_missing_presentation(self) -> bool {
        matches!(self, Self::MissingPresentation)
    }
}

/// Finite resource ceilings for one OLE1 object byte sequence.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum complete object byte length, including all headers.
    pub max_bytes: usize,
    /// Maximum embedded native-data payload length.
    pub max_native_bytes: usize,
    /// Maximum complete presentation byte length.
    pub max_presentation_bytes: usize,
    /// Maximum payload bytes in one length-prefixed string.
    pub max_string_bytes: usize,
    /// Maximum payload bytes in one registered clipboard name.
    pub max_registered_format_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_bytes: DEFAULT_MAX_BYTES,
            max_native_bytes: DEFAULT_MAX_BYTES,
            max_presentation_bytes: DEFAULT_MAX_BYTES,
            max_string_bytes: DEFAULT_MAX_STRING_BYTES,
            max_registered_format_bytes: DEFAULT_MAX_REGISTERED_BYTES,
        }
    }
}

impl Limits {
    fn validate(self) -> Result<Self, OleError> {
        let values = [
            ("OLE1 bytes", self.max_bytes),
            ("OLE1 native bytes", self.max_native_bytes),
            ("OLE1 presentation bytes", self.max_presentation_bytes),
            ("OLE1 string bytes", self.max_string_bytes),
            (
                "OLE1 registered format bytes",
                self.max_registered_format_bytes,
            ),
        ];
        for (resource, value) in values {
            if value == 0 {
                return Err(OleError::InvalidLimit {
                    resource,
                    value: 0,
                    maximum: HARD_MAX_BYTES as u64,
                });
            }
            if value > HARD_MAX_BYTES {
                return Err(OleError::InvalidLimit {
                    resource,
                    value: value as u64,
                    maximum: HARD_MAX_BYTES as u64,
                });
            }
        }
        for (resource, value, maximum) in [
            ("OLE1 native bytes", self.max_native_bytes, self.max_bytes),
            (
                "OLE1 presentation bytes",
                self.max_presentation_bytes,
                self.max_bytes,
            ),
            ("OLE1 string bytes", self.max_string_bytes, self.max_bytes),
            (
                "OLE1 registered format bytes",
                self.max_registered_format_bytes,
                self.max_string_bytes,
            ),
        ] {
            if value > maximum {
                return Err(OleError::InvalidLimit {
                    resource,
                    value: value as u64,
                    maximum: maximum as u64,
                });
            }
        }
        Ok(self)
    }
}

/// ANSI or UTF-16LE bytes used by an OLE1 length-prefixed string.
///
/// The bytes contain the terminating NUL when the string is non-empty, but do
/// not contain the four-byte length prefix.  Keeping this representation raw
/// avoids a code-page guess and allows source spelling to survive unchanged.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncodedString {
    encoding: TextEncoding,
    payload: Arc<[u8]>,
}

impl EncodedString {
    /// Creates a validated ANSI string payload, including its terminating NUL.
    pub fn ansi(bytes: impl Into<Vec<u8>>) -> Result<Self, OleError> {
        Self::from_payload(TextEncoding::Ansi, bytes.into())
    }

    /// Creates a validated UTF-16LE string payload, including its terminating
    /// UTF-16 NUL code unit.
    pub fn unicode(bytes: impl Into<Vec<u8>>) -> Result<Self, OleError> {
        Self::from_payload(TextEncoding::Unicode, bytes.into())
    }

    fn from_payload(encoding: TextEncoding, bytes: Vec<u8>) -> Result<Self, OleError> {
        validate_string_payload(encoding, &bytes, DEFAULT_MAX_STRING_BYTES)?;
        Ok(Self {
            encoding,
            payload: bytes.into(),
        })
    }

    /// Returns the encoding profile of this string.
    #[must_use]
    pub const fn encoding(&self) -> TextEncoding {
        self.encoding
    }

    /// Returns the payload bytes, including the encoded terminating NUL.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    fn encoded_len(&self) -> Result<usize, OleError> {
        self.payload
            .len()
            .checked_add(4)
            .ok_or_else(|| invalid("OLE1 string encoded length overflows"))
    }
}

/// The encoding used by one length-prefixed string.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextEncoding {
    /// ANSI bytes, with a one-byte NUL terminator.
    Ansi,
    /// UTF-16LE bytes, with a two-byte NUL terminator.
    Unicode,
}

/// A borrowed length-prefixed string with its exact encoded source bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodedStringRef<'a> {
    encoding: TextEncoding,
    encoded: &'a [u8],
    payload: &'a [u8],
}

impl<'a> EncodedStringRef<'a> {
    /// Returns the source encoding profile.
    #[must_use]
    pub const fn encoding(self) -> TextEncoding {
        self.encoding
    }

    /// Returns the exact four-byte length prefix plus payload.
    #[must_use]
    pub const fn encoded_bytes(self) -> &'a [u8] {
        self.encoded
    }

    /// Returns the payload bytes, including its terminating NUL when present.
    #[must_use]
    pub const fn payload(self) -> &'a [u8] {
        self.payload
    }

    /// Whether the source string uses the empty length form.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.payload.is_empty()
    }
}

/// The OLE1 object role declared by an ObjectHeader.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ObjectKind {
    /// The object carries an opaque native-data payload.
    Embedded,
    /// The object carries an inert external-link descriptor.
    Linked,
}

impl ObjectKind {
    const fn format_id(self) -> u32 {
        match self {
            Self::Embedded => FORMAT_ID_EMBEDDED,
            Self::Linked => FORMAT_ID_LINKED,
        }
    }
}

/// A canonical authoring ObjectHeader.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ObjectHeader {
    kind: ObjectKind,
    ole_version: u32,
    class_name: EncodedString,
    topic_name: EncodedString,
    item_name: EncodedString,
}

impl ObjectHeader {
    /// Creates an embedded ObjectHeader with the conventional ignored OLE1
    /// version value `0x00000501`.
    pub fn embedded(
        class_name: EncodedString,
        topic_name: EncodedString,
        item_name: EncodedString,
    ) -> Result<Self, OleError> {
        Self::with_version(
            ObjectKind::Embedded,
            CANONICAL_OLE_VERSION,
            class_name,
            topic_name,
            item_name,
        )
    }

    /// Creates a linked ObjectHeader with the conventional ignored OLE1
    /// version value `0x00000501`.
    pub fn linked(
        class_name: EncodedString,
        topic_name: EncodedString,
        item_name: EncodedString,
    ) -> Result<Self, OleError> {
        Self::with_version(
            ObjectKind::Linked,
            CANONICAL_OLE_VERSION,
            class_name,
            topic_name,
            item_name,
        )
    }

    /// Creates an ObjectHeader with a caller-supplied ignored OLE version.
    pub fn with_version(
        kind: ObjectKind,
        ole_version: u32,
        class_name: EncodedString,
        topic_name: EncodedString,
        item_name: EncodedString,
    ) -> Result<Self, OleError> {
        ensure_ansi(&class_name, "ObjectHeader ClassName")?;
        ensure_ansi(&topic_name, "ObjectHeader TopicName")?;
        ensure_ansi(&item_name, "ObjectHeader ItemName")?;
        ensure_nonempty_string(&class_name, "ObjectHeader ClassName")?;
        if kind == ObjectKind::Linked {
            ensure_linked_topic_name(&topic_name)?;
        }
        Ok(Self {
            kind,
            ole_version,
            class_name,
            topic_name,
            item_name,
        })
    }

    /// Returns the object role.
    #[must_use]
    pub const fn kind(&self) -> ObjectKind {
        self.kind
    }

    /// Returns the ignored OLE version field.
    #[must_use]
    pub const fn ole_version(&self) -> u32 {
        self.ole_version
    }

    /// Returns the OLE1 format identifier implied by this header.
    #[must_use]
    pub const fn format_id(&self) -> u32 {
        self.kind.format_id()
    }

    /// Returns the raw ANSI class name.
    #[must_use]
    pub fn class_name(&self) -> &EncodedString {
        &self.class_name
    }

    /// Returns the raw ANSI topic name.
    #[must_use]
    pub fn topic_name(&self) -> &EncodedString {
        &self.topic_name
    }

    /// Returns the raw ANSI item name.
    #[must_use]
    pub fn item_name(&self) -> &EncodedString {
        &self.item_name
    }
}

/// The five normative OLE1 presentation forms available to canonical
/// authoring.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Presentation {
    /// A `METAFILEPICT` presentation with opaque metafile bytes.
    MetaFile {
        /// Signed logical width from the wire.
        width: i32,
        /// Signed logical height from the wire.
        height: i32,
        /// Opaque presentation payload.
        data: Vec<u8>,
    },
    /// A `BITMAP` presentation with opaque Bitmap16 bytes.
    Bitmap {
        /// Signed width from the wire.
        width: i32,
        /// Signed height from the wire.
        height: i32,
        /// Opaque presentation payload.
        data: Vec<u8>,
    },
    /// A `DIB` presentation with opaque device-independent bitmap bytes.
    Dib {
        /// Signed width from the wire.
        width: i32,
        /// Signed height from the wire.
        height: i32,
        /// Opaque presentation payload.
        data: Vec<u8>,
    },
    /// A standard clipboard presentation.
    StandardClipboard {
        /// Generic presentation class name, encoded as ANSI bytes.
        class_name: EncodedString,
        /// Standard clipboard identifier.
        clipboard_format: u32,
        /// Opaque presentation payload.
        data: Vec<u8>,
    },
    /// A registered clipboard presentation with an explicitly selected ANSI
    /// or UTF-16LE format-name encoding.
    RegisteredClipboard {
        /// Generic presentation class name, encoded as ANSI bytes.
        class_name: EncodedString,
        /// Registered clipboard name, with explicit source encoding.
        format_name: EncodedString,
        /// Opaque presentation payload.
        data: Vec<u8>,
    },
}

impl Presentation {
    /// Creates a canonical `METAFILEPICT` presentation.
    pub fn metafile(width: i32, height: i32, data: Vec<u8>) -> Result<Self, OleError> {
        validate_new_data(&data)?;
        Ok(Self::MetaFile {
            width,
            height,
            data,
        })
    }

    /// Creates a canonical `BITMAP` presentation.
    pub fn bitmap(width: i32, height: i32, data: Vec<u8>) -> Result<Self, OleError> {
        validate_new_data(&data)?;
        Ok(Self::Bitmap {
            width,
            height,
            data,
        })
    }

    /// Creates a canonical `DIB` presentation.
    pub fn dib(width: i32, height: i32, data: Vec<u8>) -> Result<Self, OleError> {
        validate_new_data(&data)?;
        Ok(Self::Dib {
            width,
            height,
            data,
        })
    }

    /// Creates a canonical standard clipboard presentation.
    pub fn standard_clipboard(
        class_name: EncodedString,
        clipboard_format: u32,
        data: Vec<u8>,
    ) -> Result<Self, OleError> {
        ensure_generic_class(&class_name)?;
        if !is_standard_clipboard_format(clipboard_format) {
            return Err(invalid(
                "standard clipboard format is not one of the MS-OLEDS standard identifiers",
            ));
        }
        validate_new_data(&data)?;
        Ok(Self::StandardClipboard {
            class_name,
            clipboard_format,
            data,
        })
    }

    /// Creates a canonical registered clipboard presentation.  The optional
    /// source-name form is intentionally not authored; a registered name is
    /// required so the resulting bytes identify their format unambiguously.
    pub fn registered_clipboard(
        class_name: EncodedString,
        format_name: EncodedString,
        data: Vec<u8>,
    ) -> Result<Self, OleError> {
        ensure_generic_class(&class_name)?;
        ensure_registered_format_name(&format_name, DEFAULT_MAX_REGISTERED_BYTES)?;
        validate_new_data(&data)?;
        Ok(Self::RegisteredClipboard {
            class_name,
            format_name,
            data,
        })
    }
}

/// A borrowed view of one complete OLE1 object.
#[derive(Debug, Clone)]
pub struct ObjectRef<'a> {
    bytes: &'a [u8],
    layout: Layout,
}

impl<'a> ObjectRef<'a> {
    /// Parses a borrowed object under the default limits and strict profile.
    pub fn parse(bytes: &'a [u8]) -> Result<Self, OleError> {
        Self::parse_with_limits_and_compatibility(bytes, Limits::default(), Compatibility::Strict)
    }

    /// Parses a borrowed object under explicit limits.
    pub fn parse_with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self, OleError> {
        Self::parse_with_limits_and_compatibility(bytes, limits, Compatibility::Strict)
    }

    /// Parses a borrowed object under explicit limits and compatibility.
    pub fn parse_with_limits_and_compatibility(
        bytes: &'a [u8],
        limits: Limits,
        compatibility: Compatibility,
    ) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        check_limit("OLE1 bytes", bytes.len(), limits.max_bytes)?;
        let layout = parse_layout(bytes, limits, compatibility)?;
        Ok(Self { bytes, layout })
    }

    /// Returns the exact source bytes.
    #[must_use]
    pub const fn bytes(&self) -> &'a [u8] {
        self.bytes
    }

    /// Returns the object role.
    #[must_use]
    pub const fn kind(&self) -> ObjectKind {
        self.layout.kind
    }

    /// Returns the borrowed ObjectHeader projection.
    #[must_use]
    pub fn header(&self) -> ObjectHeaderRef<'a> {
        ObjectHeaderRef::from_layout(self.bytes, &self.layout.header, self.layout.kind)
    }

    /// Returns the opaque native bytes for an embedded object.
    #[must_use]
    pub fn native_data(&self) -> Option<&'a [u8]> {
        self.layout
            .native
            .as_ref()
            .map(|range| &self.bytes[range.clone()])
    }

    /// Returns the linked object's opaque network name.
    #[must_use]
    pub fn network_name(&self) -> Option<EncodedStringRef<'a>> {
        self.layout
            .network
            .as_ref()
            .map(|field| string_ref(self.bytes, field))
    }

    /// Returns the linked object's update hint.
    #[must_use]
    pub const fn link_update_option(&self) -> Option<u32> {
        self.layout.link_update
    }

    /// Returns the typed or raw presentation projection.
    #[must_use]
    pub fn presentation(&self) -> PresentationRef<'a> {
        PresentationRef {
            bytes: self.bytes,
            layout: self.layout.presentation.clone(),
        }
    }

    /// Whether this source has a complete supported typed rewrite closure.
    #[must_use]
    pub const fn rewrite_safe(&self) -> bool {
        self.layout.rewrite_safe
    }
}

/// A borrowed ObjectHeader projection.
#[derive(Debug, Clone, Copy)]
pub struct ObjectHeaderRef<'a> {
    kind: ObjectKind,
    ole_version: u32,
    format_id: u32,
    class_name: EncodedStringRef<'a>,
    topic_name: EncodedStringRef<'a>,
    item_name: EncodedStringRef<'a>,
}

impl<'a> ObjectHeaderRef<'a> {
    fn from_layout(bytes: &'a [u8], layout: &HeaderLayout, kind: ObjectKind) -> Self {
        Self {
            kind,
            ole_version: layout.ole_version,
            format_id: layout.format_id,
            class_name: string_ref(bytes, &layout.class_name),
            topic_name: string_ref(bytes, &layout.topic_name),
            item_name: string_ref(bytes, &layout.item_name),
        }
    }

    /// Returns the object role.
    #[must_use]
    pub const fn kind(self) -> ObjectKind {
        self.kind
    }

    /// Returns the ignored OLE version.
    #[must_use]
    pub const fn ole_version(self) -> u32 {
        self.ole_version
    }

    /// Returns the OLE1 format identifier.
    #[must_use]
    pub const fn format_id(self) -> u32 {
        self.format_id
    }

    /// Returns the class name.
    #[must_use]
    pub const fn class_name(self) -> EncodedStringRef<'a> {
        self.class_name
    }

    /// Returns the topic name.
    #[must_use]
    pub const fn topic_name(self) -> EncodedStringRef<'a> {
        self.topic_name
    }

    /// Returns the item name.
    #[must_use]
    pub const fn item_name(self) -> EncodedStringRef<'a> {
        self.item_name
    }
}

/// The recognized presentation form of a borrowed view.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PresentationKind {
    /// The standard `METAFILEPICT` form.
    MetaFile,
    /// The standard `BITMAP` form.
    Bitmap,
    /// The standard `DIB` form.
    Dib,
    /// A standard clipboard form.
    StandardClipboard,
    /// A registered clipboard form.
    RegisteredClipboard,
    /// A source-only FormatID-zero compatibility presentation.
    RawFormatZero,
    /// An embedded source that ends after native data under the explicit
    /// [`Compatibility::MissingPresentation`] profile.
    MissingPresentation,
}

/// The outer-form interpretation retained for a registered presentation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RegisteredForm {
    /// A registered format name is present in the optional outer field.
    WithName,
    /// The optional outer registered format name is absent.
    WithoutName,
    /// More than one complete grammar matched the exact enclosing bytes.
    Ambiguous,
}

/// A borrowed view of one OLE1 presentation.
#[derive(Debug, Clone)]
pub struct PresentationRef<'a> {
    bytes: &'a [u8],
    layout: PresentationLayout,
}

impl<'a> PresentationRef<'a> {
    /// Parses a standalone presentation object under default strict limits.
    pub fn parse(bytes: &'a [u8]) -> Result<Self, OleError> {
        Self::parse_with_limits_and_compatibility(bytes, Limits::default(), Compatibility::Strict)
    }

    /// Parses a standalone presentation object under explicit limits.
    pub fn parse_with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self, OleError> {
        Self::parse_with_limits_and_compatibility(bytes, limits, Compatibility::Strict)
    }

    /// Parses a standalone presentation object under explicit limits and
    /// compatibility.  A FormatID-zero header is valid as a standalone raw
    /// presentation; the empty-presentation compatibility is only needed by
    /// an ObjectHeader parser.
    pub fn parse_with_limits_and_compatibility(
        bytes: &'a [u8],
        limits: Limits,
        compatibility: Compatibility,
    ) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        check_limit("OLE1 presentation bytes", bytes.len(), limits.max_bytes)?;
        let layout = parse_presentation(bytes, 0, bytes.len(), limits, compatibility, false)?;
        Ok(Self { bytes, layout })
    }

    /// Returns exact presentation source bytes.
    #[must_use]
    pub fn bytes(&self) -> &'a [u8] {
        &self.bytes[self.layout.range.clone()]
    }

    /// Returns the recognized presentation form.
    #[must_use]
    pub const fn kind(&self) -> PresentationKind {
        self.layout.kind
    }

    /// Returns the presentation header's ignored OLE version.  A missing
    /// presentation has no header and reports zero.
    #[must_use]
    pub const fn ole_version(&self) -> u32 {
        self.layout.header.ole_version
    }

    /// Returns the presentation header FormatID.  A
    /// [`PresentationKind::MissingPresentation`] view has no header and
    /// reports zero; inspect [`Self::kind`] to distinguish that source-only
    /// compatibility case from a raw FormatID-zero header.
    #[must_use]
    pub const fn format_id(&self) -> u32 {
        self.layout.header.format_id
    }

    /// Returns the generic class name, when the source header has one.
    #[must_use]
    pub fn class_name(&self) -> Option<EncodedStringRef<'a>> {
        self.layout
            .header
            .class_name
            .as_ref()
            .map(|field| string_ref(self.bytes, field))
    }

    /// Returns signed width for one of the three standard forms.
    #[must_use]
    pub const fn width(&self) -> Option<i32> {
        self.layout.width
    }

    /// Returns signed height for one of the three standard forms.
    #[must_use]
    pub const fn height(&self) -> Option<i32> {
        self.layout.height
    }

    /// Returns the four raw metafile reserved words, when applicable.
    #[must_use]
    pub const fn metafile_reserved(&self) -> Option<[u16; 4]> {
        self.layout.metafile_reserved
    }

    /// Returns the standard or registered clipboard identifier.
    #[must_use]
    pub const fn clipboard_format(&self) -> Option<u32> {
        self.layout.clipboard_format
    }

    /// Returns the registered outer-form interpretation.
    #[must_use]
    pub const fn registered_form(&self) -> Option<RegisteredForm> {
        self.layout.registered_form
    }

    /// Returns the registered format name when the source grammar identified
    /// one uniquely.
    #[must_use]
    pub fn registered_format_name(&self) -> Option<EncodedStringRef<'a>> {
        self.layout
            .registered_name
            .as_ref()
            .map(|field| string_ref(self.bytes, field))
    }

    /// Returns the opaque presentation payload when its enclosing form is
    /// unambiguous.
    #[must_use]
    pub fn data(&self) -> Option<&'a [u8]> {
        self.layout
            .data
            .as_ref()
            .map(|range| &self.bytes[range.clone()])
    }

    /// Whether a changed typed rewrite is safe for this source.
    #[must_use]
    pub const fn rewrite_safe(&self) -> bool {
        self.layout.rewrite_safe
    }
}

/// A source-backed immutable OLE1 object snapshot.
#[derive(Debug, Clone)]
pub struct ObjectSnapshot {
    source: Arc<[u8]>,
    layout: Layout,
    limits: Limits,
    compatibility: Compatibility,
    revision: Revision,
}

impl ObjectSnapshot {
    /// Parses and owns one source copy under default limits.
    pub fn parse(bytes: &[u8]) -> Result<Self, OleError> {
        Self::parse_with_limits_and_compatibility(bytes, Limits::default(), Compatibility::Strict)
    }

    /// Parses and owns one source copy under explicit limits.
    pub fn parse_with_limits(bytes: &[u8], limits: Limits) -> Result<Self, OleError> {
        Self::parse_with_limits_and_compatibility(bytes, limits, Compatibility::Strict)
    }

    /// Parses and owns one source copy under explicit limits and compatibility.
    pub fn parse_with_limits_and_compatibility(
        bytes: &[u8],
        limits: Limits,
        compatibility: Compatibility,
    ) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        check_limit("OLE1 bytes", bytes.len(), limits.max_bytes)?;
        // Admit against the caller's borrowed bytes first.  In particular,
        // oversized or malformed native/presentation fields must not retain a
        // source allocation merely to discover that they are refused.
        let layout = parse_layout(bytes, limits, compatibility)?;
        let revision = Revision::of(bytes);
        let source = Arc::<[u8]>::from(bytes);
        Ok(Self {
            source,
            layout,
            limits,
            compatibility,
            revision,
        })
    }

    /// Parses an existing immutable source allocation without copying it.
    pub fn parse_shared(
        source: Arc<[u8]>,
        limits: Limits,
        compatibility: Compatibility,
    ) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        check_limit("OLE1 bytes", source.len(), limits.max_bytes)?;
        let layout = parse_layout(&source, limits, compatibility)?;
        let revision = Revision::of(&source);
        Ok(Self {
            source,
            layout,
            limits,
            compatibility,
            revision,
        })
    }

    /// Creates and validates a canonical embedded object.
    pub fn new_embedded(
        header: ObjectHeader,
        native_data: Vec<u8>,
        presentation: Presentation,
    ) -> Result<Self, OleError> {
        Self::new_embedded_with_limits(header, native_data, presentation, Limits::default())
    }

    /// Creates and validates a canonical embedded object under explicit limits.
    pub fn new_embedded_with_limits(
        header: ObjectHeader,
        native_data: Vec<u8>,
        presentation: Presentation,
        limits: Limits,
    ) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        if header.kind != ObjectKind::Embedded {
            return Err(invalid(
                "embedded authoring requires an embedded ObjectHeader",
            ));
        }
        check_limit(
            "OLE1 native bytes",
            native_data.len(),
            limits.max_native_bytes,
        )?;
        let bytes = encode_new_object(&header, Some(&native_data), None, 0, &presentation, limits)?;
        Self::parse_shared(bytes.into(), limits, Compatibility::Strict)
    }

    /// Creates and validates a canonical linked object.
    pub fn new_linked(
        header: ObjectHeader,
        network_name: EncodedString,
        link_update_option: u32,
        presentation: Presentation,
    ) -> Result<Self, OleError> {
        Self::new_linked_with_limits(
            header,
            network_name,
            link_update_option,
            presentation,
            Limits::default(),
        )
    }

    /// Creates and validates a canonical linked object under explicit limits.
    pub fn new_linked_with_limits(
        header: ObjectHeader,
        network_name: EncodedString,
        link_update_option: u32,
        presentation: Presentation,
        limits: Limits,
    ) -> Result<Self, OleError> {
        let limits = limits.validate()?;
        if header.kind != ObjectKind::Linked {
            return Err(invalid("linked authoring requires a linked ObjectHeader"));
        }
        ensure_ansi(&network_name, "LinkedObject NetworkName")?;
        let bytes = encode_new_object(
            &header,
            None,
            Some(&network_name),
            link_update_option,
            &presentation,
            limits,
        )?;
        Self::parse_shared(bytes.into(), limits, Compatibility::Strict)
    }

    /// Returns the exact immutable source bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Returns the shared source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.source)
    }

    /// Returns the object role.
    #[must_use]
    pub const fn kind(&self) -> ObjectKind {
        self.layout.kind
    }

    /// Returns the borrowed ObjectHeader projection.
    #[must_use]
    pub fn header(&self) -> ObjectHeaderRef<'_> {
        ObjectHeaderRef::from_layout(&self.source, &self.layout.header, self.layout.kind)
    }

    /// Returns embedded native data, if present.
    #[must_use]
    pub fn native_data(&self) -> Option<&[u8]> {
        self.layout
            .native
            .as_ref()
            .map(|range| &self.source[range.clone()])
    }

    /// Returns the linked network name, if present.
    #[must_use]
    pub fn network_name(&self) -> Option<EncodedStringRef<'_>> {
        self.layout
            .network
            .as_ref()
            .map(|field| string_ref(&self.source, field))
    }

    /// Returns the linked update hint, if present.
    #[must_use]
    pub const fn link_update_option(&self) -> Option<u32> {
        self.layout.link_update
    }

    /// Returns a borrowed presentation projection.
    #[must_use]
    pub fn presentation(&self) -> PresentationRef<'_> {
        PresentationRef {
            bytes: &self.source,
            layout: self.layout.presentation.clone(),
        }
    }

    /// Whether a changed typed rewrite is admitted for this source.
    #[must_use]
    pub const fn rewrite_safe(&self) -> bool {
        self.layout.rewrite_safe
    }

    /// Returns the limits retained by this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Returns the compatibility profile retained by this snapshot.
    #[must_use]
    pub const fn compatibility(&self) -> Compatibility {
        self.compatibility
    }

    /// Returns the exact source revision used by patches.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Starts an isolated source-backed edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            current: self.clone(),
        }
    }

    /// Alias for [`Self::edit`].
    #[must_use]
    pub fn transaction(&self) -> Edit {
        self.edit()
    }
}

/// A source-backed OLE1 transaction.
#[derive(Debug, Clone)]
pub struct Edit {
    source: ObjectSnapshot,
    current: ObjectSnapshot,
}

impl Edit {
    /// Returns the immutable source snapshot used by this edit.
    #[must_use]
    pub const fn source(&self) -> &ObjectSnapshot {
        &self.source
    }

    /// Returns the currently staged snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &ObjectSnapshot {
        &self.current
    }

    /// Whether staged bytes differ from the source bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.source.bytes() != self.current.bytes()
    }

    /// Replaces the ObjectHeader class name.
    pub fn set_class_name(&mut self, value: EncodedString) -> Result<bool, OleError> {
        ensure_ansi(&value, "ObjectHeader ClassName")?;
        ensure_nonempty_string(&value, "ObjectHeader ClassName")?;
        if string_equals(
            &self.current.source,
            &self.current.layout.header.class_name,
            &value,
        ) {
            return Ok(false);
        }
        self.rewrite(EditChange::ClassName(value))
    }

    /// Replaces the ObjectHeader topic name.
    pub fn set_topic_name(&mut self, value: EncodedString) -> Result<bool, OleError> {
        ensure_ansi(&value, "ObjectHeader TopicName")?;
        if self.current.kind() == ObjectKind::Linked {
            ensure_linked_topic_name(&value)?;
        }
        if string_equals(
            &self.current.source,
            &self.current.layout.header.topic_name,
            &value,
        ) {
            return Ok(false);
        }
        self.rewrite(EditChange::TopicName(value))
    }

    /// Replaces the ObjectHeader item name.
    pub fn set_item_name(&mut self, value: EncodedString) -> Result<bool, OleError> {
        ensure_ansi(&value, "ObjectHeader ItemName")?;
        if string_equals(
            &self.current.source,
            &self.current.layout.header.item_name,
            &value,
        ) {
            return Ok(false);
        }
        self.rewrite(EditChange::ItemName(value))
    }

    /// Replaces embedded native bytes.
    pub fn set_native_data(&mut self, value: Vec<u8>) -> Result<bool, OleError> {
        let Some(range) = self.current.layout.native.as_ref() else {
            return Err(unsupported("linked objects do not have native data"));
        };
        check_limit(
            "OLE1 native bytes",
            value.len(),
            self.current.limits.max_native_bytes,
        )?;
        if &self.current.source[range.clone()] == value.as_slice() {
            return Ok(false);
        }
        self.rewrite(EditChange::NativeData(value))
    }

    /// Replaces a linked object's network name.
    pub fn set_network_name(&mut self, value: EncodedString) -> Result<bool, OleError> {
        ensure_ansi(&value, "LinkedObject NetworkName")?;
        let Some(field) = self.current.layout.network.as_ref() else {
            return Err(unsupported("embedded objects do not have a network name"));
        };
        if string_equals(&self.current.source, field, &value) {
            return Ok(false);
        }
        self.rewrite(EditChange::NetworkName(value))
    }

    /// Replaces the linked update hint.
    pub fn set_link_update_option(&mut self, value: u32) -> Result<bool, OleError> {
        let Some(current) = self.current.layout.link_update else {
            return Err(unsupported(
                "embedded objects do not have a link update hint",
            ));
        };
        if current == value {
            return Ok(false);
        }
        self.rewrite(EditChange::LinkUpdateOption(value))
    }

    /// Replaces the signed standard-presentation width.
    pub fn set_width(&mut self, value: i32) -> Result<bool, OleError> {
        let Some(current) = self.current.layout.presentation.width else {
            return Err(unsupported("this presentation has no signed dimensions"));
        };
        if current == value {
            return Ok(false);
        }
        self.rewrite(EditChange::Width(value))
    }

    /// Replaces the signed standard-presentation height.
    pub fn set_height(&mut self, value: i32) -> Result<bool, OleError> {
        let Some(current) = self.current.layout.presentation.height else {
            return Err(unsupported("this presentation has no signed dimensions"));
        };
        if current == value {
            return Ok(false);
        }
        self.rewrite(EditChange::Height(value))
    }

    /// Replaces the opaque presentation payload while preserving its source
    /// form, reserved words, and encoded clipboard fields.
    pub fn set_presentation_data(&mut self, value: Vec<u8>) -> Result<bool, OleError> {
        let Some(range) = self.current.layout.presentation.data.as_ref() else {
            return Err(unsupported(
                "this presentation has no unambiguous data field",
            ));
        };
        check_limit(
            "OLE1 presentation data",
            value.len(),
            self.current.limits.max_presentation_bytes,
        )?;
        if &self.current.source[range.clone()] == value.as_slice() {
            return Ok(false);
        }
        self.rewrite(EditChange::PresentationData(value))
    }

    /// Replaces the complete presentation with one canonical five-form value.
    pub fn set_presentation(&mut self, value: Presentation) -> Result<bool, OleError> {
        self.rewrite(EditChange::Presentation(value))
    }

    /// Publishes the staged candidate and a reversible exact-source patch.
    pub fn commit(self) -> Result<Commit, OleError> {
        let changed = self.is_changed();
        let patch = Patch::new(self.source.clone(), self.current.clone());
        Ok(Commit {
            snapshot: self.current,
            patch,
            changed,
        })
    }

    /// Discards staged edits and returns the original source snapshot.
    #[must_use]
    pub fn rollback(self) -> ObjectSnapshot {
        self.source
    }

    fn rewrite(&mut self, change: EditChange) -> Result<bool, OleError> {
        if !self.current.layout.rewrite_safe {
            return Err(unsupported(
                "source contains a compatibility or noncanonical form whose dependency closure is not proved",
            ));
        }
        let bytes = render_change(&self.current, &change)?;
        if bytes.as_slice() == self.current.bytes() {
            return Ok(false);
        }
        // A transaction that returns to its original lexical source should
        // publish that source allocation, not a newly parsed equal byte
        // sequence.  This makes mutate/revert preserve source identity.
        if bytes.as_slice() == self.source.bytes() {
            self.current = self.source.clone();
            return Ok(true);
        }
        let candidate = ObjectSnapshot::parse_shared(
            bytes.into(),
            self.current.limits,
            self.current.compatibility,
        )?;
        self.current = candidate;
        Ok(true)
    }
}

/// A successful OLE1 publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: ObjectSnapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Returns the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &ObjectSnapshot {
        &self.snapshot
    }

    /// Returns the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether this publication changed any source bytes.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Consumes the publication into its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> ObjectSnapshot {
        self.snapshot
    }

    /// Consumes the publication into its patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Splits the publication into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (ObjectSnapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible exact-source OLE1 patch.
#[derive(Debug, Clone)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: ObjectSnapshot,
    after: ObjectSnapshot,
    changed: bool,
}

impl Patch {
    fn new(before: ObjectSnapshot, after: ObjectSnapshot) -> Self {
        let changed = before.bytes() != after.bytes();
        Self {
            base: before.revision,
            target: after.revision,
            before,
            after,
            changed,
        }
    }

    /// Returns the expected source revision.
    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Returns the resulting revision.
    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Returns the exact source bytes retained by this patch.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Returns the exact target bytes retained by this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Returns the patch source snapshot.
    #[must_use]
    pub const fn source(&self) -> &ObjectSnapshot {
        &self.before
    }

    /// Returns the patch target snapshot.
    #[must_use]
    pub const fn replacement(&self) -> &ObjectSnapshot {
        &self.after
    }

    /// Whether this patch is an exact no-op.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        !self.changed
    }

    /// Applies this patch only to its exact source snapshot.
    pub fn apply(&self, source: &ObjectSnapshot) -> Result<ObjectSnapshot, OleError> {
        if source.revision != self.base || source.bytes() != self.before.bytes() {
            return Err(OleError::InvalidFormat(
                "OLE1 patch source does not match its exact base".into(),
            ));
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Ok(self.after.clone())
    }

    /// Returns the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            changed: self.changed,
        }
    }
}

/// A compact identity for one exact source byte sequence.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    /// Returns the fingerprint value.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    /// Alias for [`Self::value`].
    #[must_use]
    pub const fn fingerprint(self) -> u64 {
        self.value()
    }

    fn of(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }
}

/// Convenient module-level snapshot alias.
pub type Snapshot = ObjectSnapshot;

#[derive(Debug, Clone)]
struct Layout {
    kind: ObjectKind,
    header: HeaderLayout,
    native: Option<Range<usize>>,
    network: Option<FieldLayout>,
    link_update: Option<u32>,
    presentation: PresentationLayout,
    rewrite_safe: bool,
}

#[derive(Debug, Clone)]
struct HeaderLayout {
    ole_version: u32,
    format_id: u32,
    class_name: FieldLayout,
    topic_name: FieldLayout,
    item_name: FieldLayout,
}

#[derive(Debug, Clone)]
struct FieldLayout {
    encoded: Range<usize>,
    payload: Range<usize>,
    encoding: TextEncoding,
}

#[derive(Debug, Clone)]
struct PresentationHeaderLayout {
    ole_version: u32,
    format_id: u32,
    class_name: Option<FieldLayout>,
}

#[derive(Debug, Clone)]
struct PresentationLayout {
    range: Range<usize>,
    header: PresentationHeaderLayout,
    kind: PresentationKind,
    width: Option<i32>,
    height: Option<i32>,
    metafile_reserved: Option<[u16; 4]>,
    clipboard_format: Option<u32>,
    registered_form: Option<RegisteredForm>,
    registered_name: Option<FieldLayout>,
    data: Option<Range<usize>>,
    rewrite_safe: bool,
}

#[derive(Debug, Clone)]
enum EditChange {
    ClassName(EncodedString),
    TopicName(EncodedString),
    ItemName(EncodedString),
    NativeData(Vec<u8>),
    NetworkName(EncodedString),
    LinkUpdateOption(u32),
    Width(i32),
    Height(i32),
    PresentationData(Vec<u8>),
    Presentation(Presentation),
}

fn parse_layout(
    bytes: &[u8],
    limits: Limits,
    compatibility: Compatibility,
) -> Result<Layout, OleError> {
    let mut offset = 0;
    let ole_version = read_u32(bytes, &mut offset, "ObjectHeader OLEVersion")?;
    let format_id = read_u32(bytes, &mut offset, "ObjectHeader FormatID")?;
    let kind = match format_id {
        FORMAT_ID_EMBEDDED => ObjectKind::Embedded,
        FORMAT_ID_LINKED => ObjectKind::Linked,
        _ => return Err(invalid("ObjectHeader FormatID must be 1 or 2")),
    };
    let class_name = parse_lp(
        bytes,
        &mut offset,
        TextEncoding::Ansi,
        limits.max_string_bytes,
        "ObjectHeader ClassName",
        false,
    )?;
    let topic_name = parse_lp(
        bytes,
        &mut offset,
        TextEncoding::Ansi,
        limits.max_string_bytes,
        "ObjectHeader TopicName",
        false,
    )?;
    let item_name = parse_lp(
        bytes,
        &mut offset,
        TextEncoding::Ansi,
        limits.max_string_bytes,
        "ObjectHeader ItemName",
        false,
    )?;
    let header = HeaderLayout {
        ole_version,
        format_id,
        class_name,
        topic_name,
        item_name,
    };
    let mut rewrite_safe = true;
    // Source reads remain permissive for legacy producers, but a changed
    // typed rewrite must not silently repair or preserve an unproved header
    // identity.  Fresh authoring applies the same checks eagerly below.
    if !field_has_content(bytes, &header.class_name)
        || (kind == ObjectKind::Linked && !field_is_linked_topic_name(bytes, &header.topic_name))
    {
        rewrite_safe = false;
    }
    let native = if kind == ObjectKind::Embedded {
        let size = read_u32(bytes, &mut offset, "EmbeddedObject NativeDataSize")?;
        let size = usize_from_u32(size, "EmbeddedObject NativeDataSize")?;
        check_limit("OLE1 native bytes", size, limits.max_native_bytes)?;
        let range = take_range(bytes, &mut offset, size, "EmbeddedObject NativeData")?;
        Some(range)
    } else {
        None
    };
    let network = if kind == ObjectKind::Linked {
        Some(parse_lp(
            bytes,
            &mut offset,
            TextEncoding::Ansi,
            limits.max_string_bytes,
            "LinkedObject NetworkName",
            false,
        )?)
    } else {
        None
    };
    let link_update = if kind == ObjectKind::Linked {
        let reserved = read_u32(bytes, &mut offset, "LinkedObject Reserved")?;
        if reserved != 0 {
            rewrite_safe = false;
        }
        Some(read_u32(
            bytes,
            &mut offset,
            "LinkedObject LinkUpdateOption",
        )?)
    } else {
        None
    };
    let presentation = if kind == ObjectKind::Embedded
        && compatibility.allows_missing_presentation()
        && offset == bytes.len()
    {
        missing_presentation(offset)
    } else {
        parse_presentation(bytes, offset, bytes.len(), limits, compatibility, true)?
    };
    if presentation.kind == PresentationKind::RawFormatZero {
        if !(kind == ObjectKind::Embedded
            && compatibility.allows_empty_presentation_header()
            && presentation.range.len() == 8)
        {
            return Err(invalid(
                "nested FormatID-zero presentation requires an explicit eight-byte empty-presentation compatibility profile",
            ));
        }
        rewrite_safe = false;
    }
    rewrite_safe &= presentation.rewrite_safe;
    if presentation.range.start != offset {
        return Err(invalid("OLE1 presentation offset is inconsistent"));
    }
    Ok(Layout {
        kind,
        header,
        native,
        network,
        link_update,
        presentation,
        rewrite_safe,
    })
}

fn missing_presentation(offset: usize) -> PresentationLayout {
    PresentationLayout {
        range: offset..offset,
        header: PresentationHeaderLayout {
            ole_version: 0,
            format_id: PRESENTATION_FORMAT_ID_NONE,
            class_name: None,
        },
        kind: PresentationKind::MissingPresentation,
        width: None,
        height: None,
        metafile_reserved: None,
        clipboard_format: None,
        registered_form: None,
        registered_name: None,
        data: None,
        rewrite_safe: false,
    }
}

fn parse_presentation(
    bytes: &[u8],
    start: usize,
    end: usize,
    limits: Limits,
    compatibility: Compatibility,
    nested: bool,
) -> Result<PresentationLayout, OleError> {
    if start > end || end - start > limits.max_presentation_bytes {
        return Err(limit_error(
            "OLE1 presentation bytes",
            end.saturating_sub(start),
            limits.max_presentation_bytes,
        ));
    }
    let mut offset = start;
    let ole_version = read_u32_bounded(
        bytes,
        &mut offset,
        end,
        "PresentationObjectHeader OLEVersion",
    )?;
    let format_id = read_u32_bounded(bytes, &mut offset, end, "PresentationObjectHeader FormatID")?;
    let class_name = match format_id {
        PRESENTATION_FORMAT_ID_NONE => {
            if nested && !compatibility.allows_empty_presentation_header() {
                return Err(invalid(
                    "nested FormatID-zero PresentationObjectHeader is not normative",
                ));
            }
            None
        },
        PRESENTATION_FORMAT_ID_CLASS => Some(parse_lp_bounded(
            bytes,
            &mut offset,
            end,
            TextEncoding::Ansi,
            limits.max_string_bytes,
            "PresentationObjectHeader ClassName",
            true,
        )?),
        _ => return Err(invalid("PresentationObjectHeader FormatID must be 0 or 5")),
    };
    let header = PresentationHeaderLayout {
        ole_version,
        format_id,
        class_name: class_name.clone(),
    };
    if format_id == PRESENTATION_FORMAT_ID_NONE {
        return Ok(PresentationLayout {
            range: start..end,
            header,
            kind: PresentationKind::RawFormatZero,
            width: None,
            height: None,
            metafile_reserved: None,
            clipboard_format: None,
            registered_form: None,
            registered_name: None,
            data: None,
            rewrite_safe: false,
        });
    }
    let class = class_name
        .as_ref()
        .ok_or_else(|| invalid("PresentationObjectHeader ClassName is absent"))?;
    let class_bytes = &bytes[class.payload.clone()];
    let class_bytes = class_bytes.strip_suffix(&[0]).unwrap_or(class_bytes);
    if class_bytes.is_empty() {
        return Err(invalid(
            "PresentationObjectHeader ClassName must be non-empty",
        ));
    }
    let (
        kind,
        width,
        height,
        metafile_reserved,
        clipboard_format,
        registered_form,
        registered_name,
        data,
        rewrite_safe,
    ) = if class_bytes == CLASS_METAFILEPICT {
        let width = read_i32_bounded(bytes, &mut offset, end, "MetaFilePresentationObject Width")?;
        let height =
            read_i32_bounded(bytes, &mut offset, end, "MetaFilePresentationObject Height")?;
        let size = usize_from_u32(
            read_u32_bounded(
                bytes,
                &mut offset,
                end,
                "MetaFilePresentationObject PresentationDataSize",
            )?,
            "MetaFilePresentationObject PresentationDataSize",
        )?;
        if size < 8 {
            return Err(invalid(
                "MetaFilePresentationObject size must include 8 reserved bytes",
            ));
        }
        let data_len = size - 8;
        check_limit(
            "OLE1 presentation data",
            data_len,
            limits.max_presentation_bytes,
        )?;
        let reserved = [
            read_u16_bounded(
                bytes,
                &mut offset,
                end,
                "MetaFilePresentationObject Reserved1",
            )?,
            read_u16_bounded(
                bytes,
                &mut offset,
                end,
                "MetaFilePresentationObject Reserved2",
            )?,
            read_u16_bounded(
                bytes,
                &mut offset,
                end,
                "MetaFilePresentationObject Reserved3",
            )?,
            read_u16_bounded(
                bytes,
                &mut offset,
                end,
                "MetaFilePresentationObject Reserved4",
            )?,
        ];
        let data = take_range_bounded(
            bytes,
            &mut offset,
            end,
            data_len,
            "MetaFilePresentationObject PresentationData",
        )?;
        if offset != end {
            return Err(invalid("MetaFilePresentationObject has trailing bytes"));
        }
        (
            PresentationKind::MetaFile,
            Some(width),
            Some(height),
            Some(reserved),
            None,
            None,
            None,
            Some(data),
            true,
        )
    } else if class_bytes == CLASS_BITMAP || class_bytes == CLASS_DIB {
        let width = read_i32_bounded(bytes, &mut offset, end, "StandardPresentationObject Width")?;
        let height =
            read_i32_bounded(bytes, &mut offset, end, "StandardPresentationObject Height")?;
        let data_len = usize_from_u32(
            read_u32_bounded(
                bytes,
                &mut offset,
                end,
                "StandardPresentationObject PresentationDataSize",
            )?,
            "StandardPresentationObject PresentationDataSize",
        )?;
        check_limit(
            "OLE1 presentation data",
            data_len,
            limits.max_presentation_bytes,
        )?;
        let data = take_range_bounded(
            bytes,
            &mut offset,
            end,
            data_len,
            "StandardPresentationObject PresentationData",
        )?;
        if offset != end {
            return Err(invalid("StandardPresentationObject has trailing bytes"));
        }
        (
            if class_bytes == CLASS_BITMAP {
                PresentationKind::Bitmap
            } else {
                PresentationKind::Dib
            },
            Some(width),
            Some(height),
            None,
            None,
            None,
            None,
            Some(data),
            true,
        )
    } else {
        let clipboard = read_u32_bounded(
            bytes,
            &mut offset,
            end,
            "ClipboardFormatHeader ClipboardFormat",
        )?;
        if clipboard != 0 {
            let data_len = usize_from_u32(
                read_u32_bounded(
                    bytes,
                    &mut offset,
                    end,
                    "StandardClipboardFormatPresentationObject PresentationDataSize",
                )?,
                "StandardClipboardFormatPresentationObject PresentationDataSize",
            )?;
            check_limit(
                "OLE1 presentation data",
                data_len,
                limits.max_presentation_bytes,
            )?;
            let data = take_range_bounded(
                bytes,
                &mut offset,
                end,
                data_len,
                "StandardClipboardFormatPresentationObject PresentationData",
            )?;
            if offset != end {
                return Err(invalid(
                    "StandardClipboardFormatPresentationObject has trailing bytes",
                ));
            }
            (
                PresentationKind::StandardClipboard,
                None,
                None,
                None,
                Some(clipboard),
                None,
                None,
                Some(data),
                // Unknown source identifiers are retained for exact reads,
                // but canonical authoring only admits the identifiers listed
                // by MS-OLEDS 2.1.1.
                is_standard_clipboard_format(clipboard),
            )
        } else {
            let candidates = registered_candidates(bytes, offset, end, limits)?;
            match candidates {
                RegisteredCandidateSet::One(candidate) => {
                    let rewrite_safe = candidate.form == RegisteredForm::WithName
                        && candidate.name.as_ref().is_some_and(|name| {
                            field_has_content(bytes, name)
                                && field_has_registered_prefix(bytes, name)
                        });
                    (
                        PresentationKind::RegisteredClipboard,
                        None,
                        None,
                        None,
                        Some(0),
                        Some(candidate.form),
                        candidate.name,
                        candidate.data,
                        rewrite_safe,
                    )
                },
                RegisteredCandidateSet::Ambiguous => (
                    PresentationKind::RegisteredClipboard,
                    None,
                    None,
                    None,
                    Some(0),
                    Some(RegisteredForm::Ambiguous),
                    None,
                    None,
                    false,
                ),
                RegisteredCandidateSet::None => {
                    return Err(invalid(
                        "RegisteredClipboardFormatPresentationObject has no complete optional-name grammar",
                    ));
                },
            }
        }
    };
    Ok(PresentationLayout {
        range: start..end,
        header,
        kind,
        width,
        height,
        metafile_reserved,
        clipboard_format,
        registered_form,
        registered_name,
        data,
        rewrite_safe,
    })
}

#[derive(Debug, Clone)]
struct RegisteredCandidate {
    form: RegisteredForm,
    name: Option<FieldLayout>,
    data: Option<Range<usize>>,
}

#[derive(Debug, Clone)]
enum RegisteredCandidateSet {
    None,
    One(RegisteredCandidate),
    Ambiguous,
}

fn registered_candidates(
    bytes: &[u8],
    start: usize,
    end: usize,
    limits: Limits,
) -> Result<RegisteredCandidateSet, OleError> {
    let without_name = try_registered_without_name(bytes, start, end, limits);
    let ansi = try_registered_with_name(bytes, start, end, limits, TextEncoding::Ansi);
    let unicode = try_registered_with_name(bytes, start, end, limits, TextEncoding::Unicode);
    let mut count = 0;
    let mut selected = None;
    for candidate in [without_name, ansi, unicode].into_iter().flatten() {
        count += 1;
        if count == 1 {
            selected = Some(candidate);
        }
    }
    Ok(match (count, selected) {
        (0, _) => RegisteredCandidateSet::None,
        (1, Some(candidate)) => RegisteredCandidateSet::One(candidate),
        (_, _) => RegisteredCandidateSet::Ambiguous,
    })
}

fn try_registered_without_name(
    bytes: &[u8],
    start: usize,
    end: usize,
    limits: Limits,
) -> Option<RegisteredCandidate> {
    let mut offset = start;
    let data_len = usize_from_u32(
        read_u32_bounded(bytes, &mut offset, end, "registered data size").ok()?,
        "registered data size",
    )
    .ok()?;
    if data_len > limits.max_presentation_bytes {
        return None;
    }
    let data = take_range_bounded(bytes, &mut offset, end, data_len, "registered data").ok()?;
    (offset == end).then_some(RegisteredCandidate {
        form: RegisteredForm::WithoutName,
        name: None,
        data: Some(data),
    })
}

fn try_registered_with_name(
    bytes: &[u8],
    start: usize,
    end: usize,
    limits: Limits,
    encoding: TextEncoding,
) -> Option<RegisteredCandidate> {
    let mut offset = start;
    let name_size = usize_from_u32(
        read_u32_bounded(bytes, &mut offset, end, "StringFormatDataSize").ok()?,
        "StringFormatDataSize",
    )
    .ok()?;
    if name_size == 0 || name_size > limits.max_string_bytes {
        return None;
    }
    let name_start = offset;
    let name_end = name_start.checked_add(name_size)?;
    if name_end > end {
        return None;
    }
    let mut name_cursor = name_start;
    let name = parse_lp_bounded(
        bytes,
        &mut name_cursor,
        name_end,
        encoding,
        limits.max_registered_format_bytes,
        "StringFormatData",
        false,
    )
    .ok()?;
    if name_cursor != name_end {
        return None;
    }
    offset = name_end;
    let data_len = usize_from_u32(
        read_u32_bounded(bytes, &mut offset, end, "registered data size").ok()?,
        "registered data size",
    )
    .ok()?;
    if data_len > limits.max_presentation_bytes {
        return None;
    }
    let data = take_range_bounded(bytes, &mut offset, end, data_len, "registered data").ok()?;
    (offset == end).then_some(RegisteredCandidate {
        form: RegisteredForm::WithName,
        name: Some(name),
        data: Some(data),
    })
}

fn render_change(snapshot: &ObjectSnapshot, change: &EditChange) -> Result<Vec<u8>, OleError> {
    let layout = &snapshot.layout;
    let source = &snapshot.source;
    let limits = snapshot.limits;
    let presentation_replacement = match change {
        EditChange::Presentation(value) => Some(value),
        _ => None,
    };
    let presentation_len = if let Some(value) = presentation_replacement {
        measure_presentation(value, limits)?
    } else {
        measure_current_presentation(snapshot, change)?
    };
    check_limit(
        "OLE1 presentation bytes",
        presentation_len,
        limits.max_presentation_bytes,
    )?;
    let current_presentation_len = layout.presentation.range.len();
    let mut total = source.len();
    total = total
        .checked_sub(current_presentation_len)
        .and_then(|value| value.checked_add(presentation_len))
        .ok_or_else(|| invalid("OLE1 output length overflows"))?;
    match change {
        EditChange::ClassName(value) => adjust_len(
            &mut total,
            layout.header.class_name.encoded.len(),
            value.encoded_len()?,
        )?,
        EditChange::TopicName(value) => adjust_len(
            &mut total,
            layout.header.topic_name.encoded.len(),
            value.encoded_len()?,
        )?,
        EditChange::ItemName(value) => adjust_len(
            &mut total,
            layout.header.item_name.encoded.len(),
            value.encoded_len()?,
        )?,
        EditChange::NativeData(value) => {
            let old = layout.native.as_ref().map_or(0, Range::len);
            adjust_len(&mut total, old, value.len())?;
        },
        EditChange::NetworkName(value) => adjust_len(
            &mut total,
            layout
                .network
                .as_ref()
                .map_or(0, |field| field.encoded.len()),
            value.encoded_len()?,
        )?,
        EditChange::LinkUpdateOption(_)
        | EditChange::Width(_)
        | EditChange::Height(_)
        | EditChange::PresentationData(_)
        | EditChange::Presentation(_) => {},
    }
    check_limit("OLE1 output bytes", total, limits.max_bytes)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(total)
        .map_err(|source| OleError::Allocation {
            resource: "OLE1 output bytes",
            source,
        })?;
    write_u32(&mut output, layout.header.ole_version);
    write_u32(&mut output, layout.header.format_id);
    write_field(
        &mut output,
        source,
        &layout.header.class_name,
        match change {
            EditChange::ClassName(value) => Some(value),
            _ => None,
        },
    )?;
    write_field(
        &mut output,
        source,
        &layout.header.topic_name,
        match change {
            EditChange::TopicName(value) => Some(value),
            _ => None,
        },
    )?;
    write_field(
        &mut output,
        source,
        &layout.header.item_name,
        match change {
            EditChange::ItemName(value) => Some(value),
            _ => None,
        },
    )?;
    match layout.kind {
        ObjectKind::Embedded => {
            let native = match change {
                EditChange::NativeData(value) => value.as_slice(),
                _ => layout
                    .native
                    .as_ref()
                    .map_or(&[][..], |range| &source[range.clone()]),
            };
            write_u32(
                &mut output,
                ensure_u32(native.len(), "EmbeddedObject NativeDataSize")?,
            );
            output.extend_from_slice(native);
        },
        ObjectKind::Linked => {
            let network = layout
                .network
                .as_ref()
                .ok_or_else(|| invalid("linked network field is absent"))?;
            write_field(
                &mut output,
                source,
                network,
                match change {
                    EditChange::NetworkName(value) => Some(value),
                    _ => None,
                },
            )?;
            write_u32(&mut output, 0);
            let update = match change {
                EditChange::LinkUpdateOption(value) => *value,
                _ => layout
                    .link_update
                    .ok_or_else(|| invalid("linked update field is absent"))?,
            };
            write_u32(&mut output, update);
        },
    }
    if let Some(value) = presentation_replacement {
        write_presentation(&mut output, value)?;
    } else {
        write_current_presentation(&mut output, snapshot, change)?;
    }
    debug_assert_eq!(output.len(), total);
    Ok(output)
}

fn measure_current_presentation(
    snapshot: &ObjectSnapshot,
    change: &EditChange,
) -> Result<usize, OleError> {
    let layout = &snapshot.layout.presentation;
    let old_data = layout.data.as_ref().map_or(0, Range::len);
    let new_data = match change {
        EditChange::PresentationData(value) => value.len(),
        _ => old_data,
    };
    let mut total = layout.range.len();
    if layout.data.is_some() {
        total = total
            .checked_sub(old_data)
            .and_then(|value| value.checked_add(new_data))
            .ok_or_else(|| invalid("OLE1 presentation length overflows"))?;
    }
    Ok(total)
}

fn write_current_presentation(
    output: &mut Vec<u8>,
    snapshot: &ObjectSnapshot,
    change: &EditChange,
) -> Result<(), OleError> {
    let source = &snapshot.source;
    let layout = &snapshot.layout.presentation;
    write_u32(output, layout.header.ole_version);
    write_u32(output, layout.header.format_id);
    if let Some(class) = layout.header.class_name.as_ref() {
        output.extend_from_slice(&source[class.encoded.clone()]);
    }
    match layout.kind {
        PresentationKind::MetaFile => {
            write_i32(
                output,
                value_or(layout.width, change, EditChangeKind::Width)?,
            );
            write_i32(
                output,
                value_or(layout.height, change, EditChangeKind::Height)?,
            );
            let data = changed_data(source, layout, change);
            write_u32(
                output,
                ensure_u32(
                    data.len()
                        .checked_add(8)
                        .ok_or_else(|| invalid("metafile presentation size overflows"))?,
                    "MetaFilePresentationObject PresentationDataSize",
                )?,
            );
            let reserved = layout
                .metafile_reserved
                .ok_or_else(|| invalid("metafile reserved words are absent"))?;
            for word in reserved {
                output.extend_from_slice(&word.to_le_bytes());
            }
            output.extend_from_slice(data);
        },
        PresentationKind::Bitmap | PresentationKind::Dib => {
            write_i32(
                output,
                value_or(layout.width, change, EditChangeKind::Width)?,
            );
            write_i32(
                output,
                value_or(layout.height, change, EditChangeKind::Height)?,
            );
            let data = changed_data(source, layout, change);
            write_u32(
                output,
                ensure_u32(
                    data.len(),
                    "StandardPresentationObject PresentationDataSize",
                )?,
            );
            output.extend_from_slice(data);
        },
        PresentationKind::StandardClipboard => {
            write_u32(
                output,
                layout
                    .clipboard_format
                    .ok_or_else(|| invalid("standard clipboard format is absent"))?,
            );
            let data = changed_data(source, layout, change);
            write_u32(
                output,
                ensure_u32(
                    data.len(),
                    "StandardClipboardFormatPresentationObject PresentationDataSize",
                )?,
            );
            output.extend_from_slice(data);
        },
        PresentationKind::RegisteredClipboard => {
            if layout.registered_form != Some(RegisteredForm::WithName) {
                return Err(unsupported(
                    "registered source form is not uniquely name-bearing",
                ));
            }
            write_u32(output, 0);
            let name = layout
                .registered_name
                .as_ref()
                .ok_or_else(|| invalid("registered format name is absent"))?;
            write_u32(
                output,
                ensure_u32(name.encoded.len(), "StringFormatDataSize")?,
            );
            output.extend_from_slice(&source[name.encoded.clone()]);
            let data = changed_data(source, layout, change);
            write_u32(
                output,
                ensure_u32(data.len(), "registered PresentationDataSize")?,
            );
            output.extend_from_slice(data);
        },
        PresentationKind::RawFormatZero => {
            return Err(unsupported(
                "FormatID-zero compatibility presentation is source-only",
            ));
        },
        PresentationKind::MissingPresentation => {
            return Err(unsupported("missing embedded presentation is source-only"));
        },
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum EditChangeKind {
    Width,
    Height,
}

fn value_or(
    current: Option<i32>,
    change: &EditChange,
    kind: EditChangeKind,
) -> Result<i32, OleError> {
    match (kind, change) {
        (EditChangeKind::Width, EditChange::Width(value)) => Ok(*value),
        (EditChangeKind::Height, EditChange::Height(value)) => Ok(*value),
        _ => current.ok_or_else(|| invalid("standard presentation dimension is absent")),
    }
}

fn changed_data<'a>(
    source: &'a [u8],
    layout: &'a PresentationLayout,
    change: &'a EditChange,
) -> &'a [u8] {
    match change {
        EditChange::PresentationData(value) => value,
        _ => layout
            .data
            .as_ref()
            .map_or(&[][..], |range| &source[range.clone()]),
    }
}

fn measure_presentation(value: &Presentation, limits: Limits) -> Result<usize, OleError> {
    let class_len = match value {
        Presentation::MetaFile { .. } => {
            check_limit(
                "OLE1 string bytes",
                CLASS_METAFILEPICT.len() + 1,
                limits.max_string_bytes,
            )?;
            canonical_class(CLASS_METAFILEPICT)?.encoded_len()?
        },
        Presentation::Bitmap { .. } => {
            check_limit(
                "OLE1 string bytes",
                CLASS_BITMAP.len() + 1,
                limits.max_string_bytes,
            )?;
            canonical_class(CLASS_BITMAP)?.encoded_len()?
        },
        Presentation::Dib { .. } => {
            check_limit(
                "OLE1 string bytes",
                CLASS_DIB.len() + 1,
                limits.max_string_bytes,
            )?;
            canonical_class(CLASS_DIB)?.encoded_len()?
        },
        Presentation::StandardClipboard { class_name, .. }
        | Presentation::RegisteredClipboard { class_name, .. } => {
            ensure_generic_class(class_name)?;
            validate_string_payload(
                class_name.encoding(),
                class_name.payload(),
                limits.max_string_bytes,
            )?;
            class_name.encoded_len()?
        },
    };
    let base = 8usize
        .checked_add(class_len)
        .ok_or_else(|| invalid("presentation header length overflows"))?;
    let total = match value {
        Presentation::MetaFile { data, .. } => base
            .checked_add(8)
            .and_then(|value| value.checked_add(12))
            .and_then(|value| value.checked_add(data.len()))
            .ok_or_else(|| invalid("metafile presentation length overflows"))?,
        Presentation::Bitmap { data, .. } | Presentation::Dib { data, .. } => base
            .checked_add(12)
            .and_then(|value| value.checked_add(data.len()))
            .ok_or_else(|| invalid("standard presentation length overflows"))?,
        Presentation::StandardClipboard {
            clipboard_format,
            data,
            ..
        } => {
            if !is_standard_clipboard_format(*clipboard_format) {
                return Err(invalid(
                    "standard clipboard format is not one of the MS-OLEDS standard identifiers",
                ));
            }
            base.checked_add(8)
                .and_then(|value| value.checked_add(data.len()))
                .ok_or_else(|| invalid("standard clipboard presentation length overflows"))?
        },
        Presentation::RegisteredClipboard {
            format_name, data, ..
        } => {
            ensure_registered_format_name(format_name, limits.max_registered_format_bytes)?;
            base.checked_add(4)
                .and_then(|value| value.checked_add(4))
                .and_then(|value| value.checked_add(format_name.encoded_len().ok()?))
                .and_then(|value| value.checked_add(4))
                .and_then(|value| value.checked_add(data.len()))
                .ok_or_else(|| invalid("registered presentation length overflows"))?
        },
    };
    check_limit(
        "OLE1 presentation bytes",
        total,
        limits.max_presentation_bytes,
    )?;
    Ok(total)
}

fn write_presentation(output: &mut Vec<u8>, value: &Presentation) -> Result<(), OleError> {
    let (class_name, dimensions, data) = match value {
        Presentation::MetaFile {
            width,
            height,
            data,
        } => (
            canonical_class(CLASS_METAFILEPICT)?,
            Some((*width, *height)),
            data.as_slice(),
        ),
        Presentation::Bitmap {
            width,
            height,
            data,
        } => (
            canonical_class(CLASS_BITMAP)?,
            Some((*width, *height)),
            data.as_slice(),
        ),
        Presentation::Dib {
            width,
            height,
            data,
        } => (
            canonical_class(CLASS_DIB)?,
            Some((*width, *height)),
            data.as_slice(),
        ),
        Presentation::StandardClipboard {
            class_name, data, ..
        }
        | Presentation::RegisteredClipboard {
            class_name, data, ..
        } => (class_name.clone(), None, data.as_slice()),
    };
    write_u32(output, CANONICAL_OLE_VERSION);
    write_u32(output, PRESENTATION_FORMAT_ID_CLASS);
    write_field_owned(output, &class_name)?;
    match value {
        Presentation::MetaFile { .. } => {
            let (width, height) =
                dimensions.ok_or_else(|| invalid("metafile dimensions are absent"))?;
            write_i32(output, width);
            write_i32(output, height);
            write_u32(
                output,
                ensure_u32(
                    data.len()
                        .checked_add(8)
                        .ok_or_else(|| invalid("metafile size overflows"))?,
                    "MetaFilePresentationObject PresentationDataSize",
                )?,
            );
            output.extend_from_slice(&[0; 8]);
            output.extend_from_slice(data);
        },
        Presentation::Bitmap { .. } | Presentation::Dib { .. } => {
            let (width, height) =
                dimensions.ok_or_else(|| invalid("standard dimensions are absent"))?;
            write_i32(output, width);
            write_i32(output, height);
            write_u32(
                output,
                ensure_u32(
                    data.len(),
                    "StandardPresentationObject PresentationDataSize",
                )?,
            );
            output.extend_from_slice(data);
        },
        Presentation::StandardClipboard {
            clipboard_format, ..
        } => {
            write_u32(output, *clipboard_format);
            write_u32(
                output,
                ensure_u32(
                    data.len(),
                    "StandardClipboardFormatPresentationObject PresentationDataSize",
                )?,
            );
            output.extend_from_slice(data);
        },
        Presentation::RegisteredClipboard { format_name, .. } => {
            write_u32(output, 0);
            write_u32(
                output,
                ensure_u32(format_name.encoded_len()?, "StringFormatDataSize")?,
            );
            write_field_owned(output, format_name)?;
            write_u32(
                output,
                ensure_u32(
                    data.len(),
                    "RegisteredClipboardFormatPresentationObject PresentationDataSize",
                )?,
            );
            output.extend_from_slice(data);
        },
    }
    Ok(())
}

fn encode_new_object(
    header: &ObjectHeader,
    native: Option<&[u8]>,
    network: Option<&EncodedString>,
    link_update_option: u32,
    presentation: &Presentation,
    limits: Limits,
) -> Result<Vec<u8>, OleError> {
    preflight_object_strings(header, network, limits)?;
    let presentation_len = measure_presentation(presentation, limits)?;
    let header_len = 8usize
        .checked_add(header.class_name.encoded_len()?)
        .and_then(|value| value.checked_add(header.topic_name.encoded_len().ok()?))
        .and_then(|value| value.checked_add(header.item_name.encoded_len().ok()?))
        .ok_or_else(|| invalid("ObjectHeader length overflows"))?;
    let body_len = match header.kind {
        ObjectKind::Embedded => {
            let native = native.ok_or_else(|| invalid("embedded native data is absent"))?;
            check_limit("OLE1 native bytes", native.len(), limits.max_native_bytes)?;
            4usize
                .checked_add(native.len())
                .ok_or_else(|| invalid("EmbeddedObject length overflows"))?
        },
        ObjectKind::Linked => {
            let network = network.ok_or_else(|| invalid("linked network name is absent"))?;
            ensure_ansi(network, "LinkedObject NetworkName")?;
            network
                .encoded_len()?
                .checked_add(8)
                .ok_or_else(|| invalid("LinkedObject length overflows"))?
        },
    };
    let total = header_len
        .checked_add(body_len)
        .and_then(|value| value.checked_add(presentation_len))
        .ok_or_else(|| invalid("OLE1 object length overflows"))?;
    check_limit("OLE1 output bytes", total, limits.max_bytes)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(total)
        .map_err(|source| OleError::Allocation {
            resource: "OLE1 output bytes",
            source,
        })?;
    write_u32(&mut output, header.ole_version);
    write_u32(&mut output, header.kind.format_id());
    write_field_owned(&mut output, &header.class_name)?;
    write_field_owned(&mut output, &header.topic_name)?;
    write_field_owned(&mut output, &header.item_name)?;
    match header.kind {
        ObjectKind::Embedded => {
            let native = native.ok_or_else(|| invalid("embedded native data is absent"))?;
            write_u32(
                &mut output,
                ensure_u32(native.len(), "EmbeddedObject NativeDataSize")?,
            );
            output.extend_from_slice(native);
        },
        ObjectKind::Linked => {
            write_field_owned(
                &mut output,
                network.ok_or_else(|| invalid("linked network name is absent"))?,
            )?;
            write_u32(&mut output, 0);
            write_u32(&mut output, link_update_option);
        },
    }
    write_presentation(&mut output, presentation)?;
    debug_assert_eq!(output.len(), total);
    Ok(output)
}

fn preflight_object_strings(
    header: &ObjectHeader,
    network: Option<&EncodedString>,
    limits: Limits,
) -> Result<(), OleError> {
    for value in [&header.class_name, &header.topic_name, &header.item_name] {
        validate_string_payload(value.encoding(), value.payload(), limits.max_string_bytes)?;
    }
    if let Some(network) = network {
        ensure_ansi(network, "LinkedObject NetworkName")?;
        validate_string_payload(
            network.encoding(),
            network.payload(),
            limits.max_string_bytes,
        )?;
    }
    Ok(())
}

fn write_field(
    output: &mut Vec<u8>,
    source: &[u8],
    field: &FieldLayout,
    replacement: Option<&EncodedString>,
) -> Result<(), OleError> {
    if let Some(value) = replacement {
        write_field_owned(output, value)
    } else {
        output.extend_from_slice(&source[field.encoded.clone()]);
        Ok(())
    }
}

fn write_field_owned(output: &mut Vec<u8>, value: &EncodedString) -> Result<(), OleError> {
    output.extend_from_slice(&ensure_u32(value.payload.len(), "OLE1 string length")?.to_le_bytes());
    output.extend_from_slice(&value.payload);
    Ok(())
}

fn string_ref<'a>(source: &'a [u8], field: &FieldLayout) -> EncodedStringRef<'a> {
    EncodedStringRef {
        encoding: field.encoding,
        encoded: &source[field.encoded.clone()],
        payload: &source[field.payload.clone()],
    }
}

fn string_equals(source: &[u8], field: &FieldLayout, value: &EncodedString) -> bool {
    field.encoding == value.encoding && &source[field.payload.clone()] == value.payload.as_ref()
}

fn parse_lp(
    bytes: &[u8],
    offset: &mut usize,
    encoding: TextEncoding,
    max: usize,
    field: &str,
    required: bool,
) -> Result<FieldLayout, OleError> {
    parse_lp_bounded(bytes, offset, bytes.len(), encoding, max, field, required)
}

fn parse_lp_bounded(
    bytes: &[u8],
    offset: &mut usize,
    end: usize,
    encoding: TextEncoding,
    max: usize,
    field: &str,
    required: bool,
) -> Result<FieldLayout, OleError> {
    let start = *offset;
    let length = usize_from_u32(read_u32_bounded(bytes, offset, end, field)?, field)?;
    if length == 0 {
        if required {
            return Err(invalid(format!("{field} must be present")));
        }
        return Ok(FieldLayout {
            encoded: start..*offset,
            payload: *offset..*offset,
            encoding,
        });
    }
    check_limit("OLE1 string bytes", length, max)?;
    let payload_start = *offset;
    let payload_range = take_range_bounded(bytes, offset, end, length, field)?;
    validate_string_payload(encoding, &bytes[payload_range.clone()], max)?;
    Ok(FieldLayout {
        encoded: start..*offset,
        payload: payload_start..*offset,
        encoding,
    })
}

fn validate_string_payload(
    encoding: TextEncoding,
    payload: &[u8],
    max: usize,
) -> Result<(), OleError> {
    check_limit("OLE1 string bytes", payload.len(), max)?;
    if payload.is_empty() {
        return Ok(());
    }
    match encoding {
        TextEncoding::Ansi => {
            if payload.last() != Some(&0) || payload[..payload.len() - 1].contains(&0) {
                return Err(invalid(
                    "OLE1 ANSI string is not a single NUL-terminated value",
                ));
            }
        },
        TextEncoding::Unicode => {
            if payload.len() % 2 != 0 || payload.len() < 2 || payload[payload.len() - 2..] != [0, 0]
            {
                return Err(invalid(
                    "OLE1 Unicode string is not an even UTF-16LE NUL-terminated value",
                ));
            }
            for pair in payload[..payload.len() - 2].chunks_exact(2) {
                if pair == [0, 0] {
                    return Err(invalid("OLE1 Unicode string contains an embedded NUL"));
                }
            }
        },
    }
    Ok(())
}

fn ensure_ansi(value: &EncodedString, field: &str) -> Result<(), OleError> {
    if value.encoding != TextEncoding::Ansi {
        return Err(invalid(format!("{field} must use ANSI encoding")));
    }
    Ok(())
}

fn ensure_nonempty_string(value: &EncodedString, field: &str) -> Result<(), OleError> {
    if string_content(value.encoding, &value.payload).is_empty() {
        return Err(invalid(format!("{field} must contain a value")));
    }
    Ok(())
}

fn ensure_linked_topic_name(value: &EncodedString) -> Result<(), OleError> {
    ensure_ansi(value, "ObjectHeader TopicName")?;
    if !is_absolute_link_topic(string_content(value.encoding, &value.payload)) {
        return Err(invalid(
            "linked ObjectHeader TopicName must be a drive-absolute or UNC path",
        ));
    }
    Ok(())
}

fn is_absolute_link_topic(value: &[u8]) -> bool {
    let drive_absolute = value.len() >= 3
        && value[0].is_ascii_alphabetic()
        && value[1] == b':'
        && matches!(value[2], b'\\' | b'/');
    let unc = value.starts_with(b"\\\\")
        && value[2..]
            .iter()
            .position(|byte| *byte == b'\\')
            .is_some_and(|separator| {
                let share_start = 2 + separator + 1;
                share_start < value.len()
                    && separator > 0
                    && !matches!(value[share_start], b'\\' | b'/')
            });
    drive_absolute || unc
}

fn ensure_registered_format_name(value: &EncodedString, max_bytes: usize) -> Result<(), OleError> {
    validate_string_payload(value.encoding, &value.payload, max_bytes)?;
    if string_content(value.encoding, &value.payload).is_empty() {
        return Err(invalid("registered clipboard name must be non-empty"));
    }
    if !registered_prefix(value.encoding, &value.payload) {
        return Err(invalid(
            "registered clipboard name must start with the MS-OLEDS OleExternal prefix",
        ));
    }
    Ok(())
}

fn string_content(encoding: TextEncoding, payload: &[u8]) -> &[u8] {
    match encoding {
        TextEncoding::Ansi => payload.strip_suffix(&[0]).unwrap_or(payload),
        TextEncoding::Unicode => {
            if payload.len() >= 2 && payload.ends_with(&[0, 0]) {
                &payload[..payload.len() - 2]
            } else {
                payload
            }
        },
    }
}

fn registered_prefix(encoding: TextEncoding, payload: &[u8]) -> bool {
    const PREFIX: &[u8] = b"OleExternal";
    let content = string_content(encoding, payload);
    match encoding {
        TextEncoding::Ansi => content.starts_with(PREFIX),
        TextEncoding::Unicode => {
            let required = PREFIX.len().saturating_mul(2);
            content.len() >= required
                && PREFIX
                    .iter()
                    .enumerate()
                    .all(|(index, byte)| content[index * 2] == *byte && content[index * 2 + 1] == 0)
        },
    }
}

fn field_has_content(bytes: &[u8], field: &FieldLayout) -> bool {
    !string_content(field.encoding, &bytes[field.payload.clone()]).is_empty()
}

fn field_is_linked_topic_name(bytes: &[u8], field: &FieldLayout) -> bool {
    field.encoding == TextEncoding::Ansi
        && is_absolute_link_topic(string_content(
            field.encoding,
            &bytes[field.payload.clone()],
        ))
}

fn field_has_registered_prefix(bytes: &[u8], field: &FieldLayout) -> bool {
    registered_prefix(field.encoding, &bytes[field.payload.clone()])
}

fn is_standard_clipboard_format(value: u32) -> bool {
    matches!(value, CF_BITMAP | CF_METAFILEPICT | CF_DIB | CF_ENHMETAFILE)
}

fn canonical_class(value: &[u8]) -> Result<EncodedString, OleError> {
    let mut payload = Vec::with_capacity(value.len() + 1);
    payload.extend_from_slice(value);
    payload.push(0);
    EncodedString::ansi(payload)
}

fn ensure_generic_class(value: &EncodedString) -> Result<(), OleError> {
    ensure_ansi(value, "generic presentation ClassName")?;
    let bytes = value.payload.strip_suffix(&[0]).unwrap_or(&value.payload);
    if bytes.is_empty()
        || bytes == CLASS_METAFILEPICT
        || bytes == CLASS_BITMAP
        || bytes == CLASS_DIB
    {
        return Err(invalid(
            "generic presentation ClassName is reserved or empty",
        ));
    }
    Ok(())
}

fn validate_new_data(data: &[u8]) -> Result<(), OleError> {
    check_limit("OLE1 presentation data", data.len(), DEFAULT_MAX_BYTES)
}

fn read_u32(bytes: &[u8], offset: &mut usize, field: &str) -> Result<u32, OleError> {
    read_u32_bounded(bytes, offset, bytes.len(), field)
}

fn read_u32_bounded(
    bytes: &[u8],
    offset: &mut usize,
    end: usize,
    field: &str,
) -> Result<u32, OleError> {
    let range = take_range_bounded(bytes, offset, end, 4, field)?;
    let raw = &bytes[range];
    Ok(u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]))
}

fn read_i32_bounded(
    bytes: &[u8],
    offset: &mut usize,
    end: usize,
    field: &str,
) -> Result<i32, OleError> {
    Ok(read_u32_bounded(bytes, offset, end, field)? as i32)
}

fn read_u16_bounded(
    bytes: &[u8],
    offset: &mut usize,
    end: usize,
    field: &str,
) -> Result<u16, OleError> {
    let range = take_range_bounded(bytes, offset, end, 2, field)?;
    let raw = &bytes[range];
    Ok(u16::from_le_bytes([raw[0], raw[1]]))
}

fn take_range(
    bytes: &[u8],
    offset: &mut usize,
    length: usize,
    field: &str,
) -> Result<Range<usize>, OleError> {
    take_range_bounded(bytes, offset, bytes.len(), length, field)
}

fn take_range_bounded(
    bytes: &[u8],
    offset: &mut usize,
    end: usize,
    length: usize,
    field: &str,
) -> Result<Range<usize>, OleError> {
    let finish = offset
        .checked_add(length)
        .ok_or_else(|| invalid(format!("{field} range overflows")))?;
    if finish > end || finish > bytes.len() {
        return Err(invalid(format!("{field} is truncated")));
    }
    let range = *offset..finish;
    *offset = finish;
    Ok(range)
}

fn usize_from_u32(value: u32, field: &str) -> Result<usize, OleError> {
    usize::try_from(value).map_err(|_| invalid(format!("{field} exceeds this platform")))
}

fn ensure_u32(value: usize, field: &str) -> Result<u32, OleError> {
    u32::try_from(value).map_err(|_| invalid(format!("{field} exceeds u32")))
}

fn write_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn write_i32(output: &mut Vec<u8>, value: i32) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn adjust_len(total: &mut usize, old: usize, new: usize) -> Result<(), OleError> {
    *total = total
        .checked_sub(old)
        .and_then(|value| value.checked_add(new))
        .ok_or_else(|| invalid("OLE1 output length overflows"))?;
    Ok(())
}

fn check_limit(resource: &'static str, observed: usize, maximum: usize) -> Result<(), OleError> {
    if observed > maximum {
        return Err(limit_error(resource, observed, maximum));
    }
    Ok(())
}

fn limit_error(resource: &'static str, observed: usize, maximum: usize) -> OleError {
    OleError::LimitExceeded {
        resource,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn invalid(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(message.into())
}

fn unsupported(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(format!("OLE1 unsupported edit: {}", message.into()))
}
