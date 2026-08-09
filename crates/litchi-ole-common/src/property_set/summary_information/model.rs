//! Semantic PIDSI values and immutable `SummaryInformation` snapshots.

use super::super::binding::Binding;
use super::super::codec::PropertySetReader;
use super::super::model::{
    CodePage, SUMMARY_INFORMATION_FMTID, Section, Stream, Value, invalid, try_clone_property_set,
    try_vec_with_capacity,
};
use chrono::{DateTime, Duration, Utc};
use litchi_cfb::{OleError, OleFile};
use std::io::{Read, Seek};

/// The `SummaryInformation` `CodePage` property identifier.
pub const CODEPAGE: u32 = super::super::model::PID_CODEPAGE;
/// The document title property identifier.
pub const TITLE: u32 = 0x0000_0002;
/// The document subject property identifier.
pub const SUBJECT: u32 = 0x0000_0003;
/// The document author property identifier.
pub const AUTHOR: u32 = 0x0000_0004;
/// The document keywords property identifier.
pub const KEYWORDS: u32 = 0x0000_0005;
/// The document comments property identifier.
pub const COMMENTS: u32 = 0x0000_0006;
/// The application-specific template property identifier.
pub const TEMPLATE: u32 = 0x0000_0007;
/// The last-author property identifier.
pub const LAST_AUTHOR: u32 = 0x0000_0008;
/// The application-specific revision-number property identifier.
pub const REVISION_NUMBER: u32 = 0x0000_0009;
/// The total editing-time FILETIME property identifier.
pub const EDIT_TIME: u32 = 0x0000_000A;
/// The most-recently-printed FILETIME property identifier.
pub const LAST_PRINTED: u32 = 0x0000_000B;
/// The document-creation FILETIME property identifier.
pub const CREATE_DTM: u32 = 0x0000_000C;
/// The most-recently-saved FILETIME property identifier.
pub const LAST_SAVE_DTM: u32 = 0x0000_000D;
/// The total page-count property identifier.
pub const PAGE_COUNT: u32 = 0x0000_000E;
/// The total word-count property identifier.
pub const WORD_COUNT: u32 = 0x0000_000F;
/// The total character-count property identifier.
pub const CHARACTER_COUNT: u32 = 0x0000_0010;
/// The optional thumbnail clipboard-data property identifier.
pub const THUMBNAIL: u32 = 0x0000_0011;
/// The creating-application property identifier.
pub const APP_NAME: u32 = 0x0000_0012;
/// The suggested document-security flag property identifier.
pub const DOC_SECURITY: u32 = 0x0000_0013;

/// Maximum UTF-8 payload accepted by a typed `SummaryInformation` string edit.
pub const MAX_TEXT_BYTES: usize = super::super::model::MAX_TYPED_TEXT_BYTES;
/// Maximum thumbnail payload accepted by the typed `SummaryInformation` owner.
pub const MAX_THUMBNAIL_BYTES: usize = 16 * 1024 * 1024;

/// A lossless 64-bit OLE FILETIME value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FileTime(u64);

impl FileTime {
    /// Creates a FILETIME from its exact wire value.
    #[must_use]
    pub const fn from_raw(raw: u64) -> Self {
        Self(raw)
    }

    /// Returns the exact 100-nanosecond wire value.
    #[must_use]
    pub const fn raw(self) -> u64 {
        self.0
    }

    /// Converts a timestamp FILETIME to UTC when it fits chrono's range.
    #[must_use]
    pub fn date_time(self) -> Option<DateTime<Utc>> {
        const EPOCH_DIFF: i64 = 116_444_736_000_000_000;
        let ticks = i64::try_from(self.0).ok()?;
        let nanos = ticks.checked_sub(EPOCH_DIFF)?.checked_mul(100)?;
        Some(DateTime::from_timestamp_nanos(nanos))
    }

    /// Converts an editing-time FILETIME to a nonnegative duration.
    #[must_use]
    pub fn duration(self) -> Option<Duration> {
        let nanos = self.0.checked_mul(100)?;
        Some(Duration::nanoseconds(i64::try_from(nanos).ok()?))
    }

    /// Encodes a UTC timestamp without losing sub-100-nanosecond precision.
    ///
    /// # Errors
    ///
    /// Returns an error if the timestamp is outside the FILETIME range or is
    /// not exactly representable in 100-nanosecond ticks.
    pub fn from_date_time(value: DateTime<Utc>) -> Result<Self, OleError> {
        const EPOCH_DIFF: i64 = 116_444_736_000_000_000;
        let nanos = value
            .timestamp_nanos_opt()
            .ok_or_else(|| invalid("DateTime is outside FILETIME precision range"))?;
        if nanos % 100 != 0 {
            return Err(invalid(
                "DateTime must be representable by 100-nanosecond FILETIME ticks",
            ));
        }
        let ticks = nanos
            .checked_div(100)
            .and_then(|ticks| ticks.checked_add(EPOCH_DIFF))
            .ok_or_else(|| invalid("FILETIME timestamp conversion overflow"))?;
        if ticks < 0 {
            return Err(invalid("DateTime cannot be represented as FILETIME"));
        }
        Ok(Self(u64::try_from(ticks).map_err(|_conversion_error| {
            invalid("DateTime cannot be represented as FILETIME")
        })?))
    }

    /// Encodes a nonnegative duration without losing sub-100-nanosecond precision.
    ///
    /// # Errors
    ///
    /// Returns an error if the duration is negative, outside the FILETIME
    /// range, or not exactly representable in 100-nanosecond ticks.
    pub fn from_duration(value: Duration) -> Result<Self, OleError> {
        let nanos = value
            .num_nanoseconds()
            .ok_or_else(|| invalid("Editing duration is outside FILETIME range"))?;
        if nanos < 0 || nanos % 100 != 0 {
            return Err(invalid(
                "Editing duration must be a nonnegative multiple of 100 nanoseconds",
            ));
        }
        Ok(Self(u64::try_from(nanos / 100).map_err(
            |_conversion_error| invalid("Editing duration cannot be represented as FILETIME"),
        )?))
    }
}

/// The clipboard tag carried by a `SummaryInformation` thumbnail.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ClipboardTag {
    /// No image data is present.
    Empty,
    /// Windows clipboard image data.
    Windows,
    /// Macintosh clipboard image data.
    Macintosh,
}

impl ClipboardTag {
    /// Decodes a standard wire clipboard tag.
    #[must_use]
    pub const fn from_raw(raw: u32) -> Option<Self> {
        match raw {
            0 => Some(Self::Empty),
            0xFFFF_FFFF => Some(Self::Windows),
            0xFFFF_FFFE => Some(Self::Macintosh),
            _ => None,
        }
    }

    /// Returns the standard wire clipboard tag.
    #[must_use]
    pub const fn raw(self) -> u32 {
        match self {
            Self::Empty => 0,
            Self::Windows => 0xFFFF_FFFF,
            Self::Macintosh => 0xFFFF_FFFE,
        }
    }
}

/// An image format permitted by the MS-OSHARED thumbnail profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ImageFormat {
    /// A METAFILEPICT payload.
    MetafilePict,
    /// An enhanced metafile payload.
    EnhancedMetafile,
    /// A JPEG payload.
    Jpeg,
}

impl ImageFormat {
    /// Decodes a standard image format identifier.
    #[must_use]
    pub const fn from_raw(raw: u32) -> Option<Self> {
        match raw {
            0x0000_0003 => Some(Self::MetafilePict),
            0x0000_000E => Some(Self::EnhancedMetafile),
            0x0000_0333 => Some(Self::Jpeg),
            _ => None,
        }
    }

    /// Returns the standard image format identifier.
    #[must_use]
    pub const fn raw(self) -> u32 {
        match self {
            Self::MetafilePict => 0x0000_0003,
            Self::EnhancedMetafile => 0x0000_000E,
            Self::Jpeg => 0x0000_0333,
        }
    }
}

/// A bounded, inert `VtThumbnailValue` payload used by `PIDSI_THUMBNAIL`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Thumbnail {
    tag: ClipboardTag,
    format: Option<ImageFormat>,
    data: Vec<u8>,
}

impl Thumbnail {
    /// Creates a bounded thumbnail after checking tag/format presence.
    ///
    /// # Errors
    ///
    /// Returns an error if the tag and image format are inconsistent or the
    /// payload exceeds the typed thumbnail size limit.
    pub fn new(
        tag: ClipboardTag,
        format: Option<ImageFormat>,
        data: Vec<u8>,
    ) -> Result<Self, OleError> {
        if tag == ClipboardTag::Empty && (format.is_some() || !data.is_empty()) {
            return Err(invalid(
                "Empty SummaryInformation thumbnails cannot carry image data",
            ));
        }
        if tag != ClipboardTag::Empty && format.is_none() {
            return Err(invalid(
                "Nonempty SummaryInformation thumbnails require an image format",
            ));
        }
        if data.len() > MAX_THUMBNAIL_BYTES {
            return Err(invalid(
                "SummaryInformation thumbnail exceeds the safety limit",
            ));
        }
        Ok(Self { tag, format, data })
    }

    /// Creates an explicit no-image thumbnail value.
    #[must_use]
    pub fn empty() -> Self {
        Self {
            tag: ClipboardTag::Empty,
            format: None,
            data: Vec::new(),
        }
    }

    /// The standard clipboard tag.
    #[must_use]
    pub const fn tag(&self) -> ClipboardTag {
        self.tag
    }

    /// The standard image format, when image data is present.
    #[must_use]
    pub const fn format(&self) -> Option<ImageFormat> {
        self.format
    }

    /// The inert thumbnail bytes.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// The thumbnail byte length.
    #[must_use]
    pub fn len(&self) -> usize {
        self.data.len()
    }

    /// Whether the thumbnail payload is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.data.is_empty()
    }

    /// Returns a zero-allocation borrowed view.
    #[must_use]
    pub fn as_ref(&self) -> ThumbnailRef<'_> {
        ThumbnailRef {
            tag: self.tag,
            format: self.format,
            data: &self.data,
        }
    }

    pub(crate) fn into_value(self) -> Result<Value, OleError> {
        let prefix: usize = if self.format.is_some() { 4 } else { 0 };
        let capacity = prefix
            .checked_add(self.data.len())
            .ok_or_else(|| invalid("SummaryInformation thumbnail size overflow"))?;
        let mut data = try_vec_with_capacity(capacity, "SummaryInformation thumbnail")?;
        if let Some(format) = self.format {
            data.extend_from_slice(&format.raw().to_le_bytes());
        }
        data.extend_from_slice(&self.data);
        Ok(Value::Clipboard {
            format: i32::from_ne_bytes(self.tag.raw().to_ne_bytes()),
            data,
        })
    }
}

/// A borrowed, zero-allocation view of a thumbnail property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ThumbnailRef<'a> {
    tag: ClipboardTag,
    format: Option<ImageFormat>,
    data: &'a [u8],
}

impl ThumbnailRef<'_> {
    pub(crate) fn from_value(value: &Value) -> Result<ThumbnailRef<'_>, OleError> {
        let Value::Clipboard {
            format: raw_format,
            data,
        } = value
        else {
            return Err(invalid("SummaryInformation Thumbnail must be VT_CF"));
        };
        let tag = ClipboardTag::from_raw(u32::from_ne_bytes(raw_format.to_ne_bytes()))
            .ok_or_else(|| invalid("SummaryInformation thumbnail has an unknown clipboard tag"))?;
        if tag == ClipboardTag::Empty {
            if !data.is_empty() {
                return Err(invalid(
                    "Empty SummaryInformation thumbnails cannot carry image data",
                ));
            }
            return Ok(ThumbnailRef {
                tag,
                format: None,
                data,
            });
        }
        if data.len() < 4 {
            return Err(invalid(
                "SummaryInformation thumbnail is missing its image format",
            ));
        }
        let format_id = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        let image_format = ImageFormat::from_raw(format_id)
            .ok_or_else(|| invalid("SummaryInformation thumbnail has an unknown image format"))?;
        Ok(ThumbnailRef {
            tag,
            format: Some(image_format),
            data: &data[4..],
        })
    }

    /// The standard clipboard tag.
    #[must_use]
    pub const fn tag(self) -> ClipboardTag {
        self.tag
    }

    /// The standard image format, when image data is present.
    #[must_use]
    pub const fn format(self) -> Option<ImageFormat> {
        self.format
    }

    /// The inert thumbnail bytes.
    #[must_use]
    pub const fn data(&self) -> &[u8] {
        self.data
    }

    /// The thumbnail byte length.
    #[must_use]
    pub const fn len(self) -> usize {
        self.data.len()
    }

    /// Whether the thumbnail payload is empty.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.data.is_empty()
    }
}

/// PIDSI document-security flags, retaining future or producer-specific bits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DocumentSecurity(u32);

impl DocumentSecurity {
    /// Password-protected document.
    pub const PASSWORD_PROTECTED: Self = Self(0x0000_0001);
    /// Read-only access is recommended.
    pub const READ_ONLY_RECOMMENDED: Self = Self(0x0000_0002);
    /// Read-only access is enforced.
    pub const READ_ONLY_ENFORCED: Self = Self(0x0000_0004);
    /// Annotation editing is locked.
    pub const LOCKED_FOR_ANNOTATIONS: Self = Self(0x0000_0008);
    /// All flags defined by [MS-OLEPS] 2.25.1.
    pub const KNOWN_BITS: Self = Self(0x0000_000F);

    /// Constructs flags while retaining unknown bits for lossless round trips.
    #[must_use]
    pub const fn new(raw: u32) -> Self {
        Self(raw)
    }

    /// Returns an empty flag set.
    #[must_use]
    pub const fn empty() -> Self {
        Self(0)
    }

    /// Returns the exact `PIDSI_DOC_SECURITY` value.
    #[must_use]
    pub const fn bits(self) -> u32 {
        self.0
    }

    /// Tests whether all bits in `flag` are present.
    #[must_use]
    pub const fn contains(self, flag: Self) -> bool {
        self.0 & flag.0 == flag.0
    }

    /// Returns bits not currently assigned by the standard.
    #[must_use]
    pub const fn unknown_bits(self) -> u32 {
        self.0 & !Self::KNOWN_BITS.0
    }

    /// Combines two checked flag sets.
    #[must_use]
    pub const fn union(self, other: Self) -> Self {
        Self(self.0 | other.0)
    }
}

impl std::ops::BitOr for DocumentSecurity {
    type Output = Self;

    fn bitor(self, rhs: Self) -> Self::Output {
        self.union(rhs)
    }
}

/// An immutable, validated `SummaryInformation` property-set view.
#[derive(Debug, Clone, PartialEq)]
pub struct Snapshot {
    pub(crate) section: Section,
}

impl Snapshot {
    /// Creates an empty `SummaryInformation` section with a required code page.
    ///
    /// # Errors
    ///
    /// Returns an error if constructing the section fails typed validation.
    pub fn new(page: CodePage) -> Result<Self, OleError> {
        let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
        section.set_page(page);
        Self::from_section(&section)
    }

    /// Validates and clones a generic `SummaryInformation` section.
    ///
    /// # Errors
    ///
    /// Returns an error if the section violates the `SummaryInformation`
    /// invariants or cannot be cloned.
    pub fn from_section(section: &Section) -> Result<Self, OleError> {
        super::validation::validate_section(section)?;
        Ok(Self {
            section: try_clone_property_set(section)?,
        })
    }

    /// Projects `SummaryInformation` from a version-zero property-set stream.
    ///
    /// # Errors
    ///
    /// Returns an error if the stream version is unsupported, the required
    /// section is absent, or the section fails typed validation.
    pub fn from_stream(stream: &Stream) -> Result<Self, OleError> {
        if stream.version != Stream::VERSION_0 {
            return Err(invalid(
                "SummaryInformation requires Property Set version 0",
            ));
        }
        let section = stream
            .section(SUMMARY_INFORMATION_FMTID)
            .ok_or_else(|| invalid("Property Set stream has no SummaryInformation section"))?;
        Self::from_section(section)
    }

    /// Reads the standard `SummaryInformation` stream from an opened CFB.
    ///
    /// # Errors
    ///
    /// Returns an error if reading or parsing the stream fails, or its section
    /// does not satisfy the typed invariants.
    pub fn from_ole<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Self, OleError> {
        let stream = ole.property_set(Binding::SummaryInformation)?;
        Self::from_stream(&stream)
    }

    /// Borrows the complete generic section, including opaque properties.
    #[must_use]
    pub const fn section(&self) -> &Section {
        &self.section
    }

    /// Returns the declared section code page.
    #[must_use]
    pub const fn codepage(&self) -> Option<CodePage> {
        self.section.page()
    }

    /// Returns a raw property for extension-specific inspection.
    #[must_use]
    pub fn property(&self, identifier: u32) -> Option<&Value> {
        self.section.property(identifier)
    }

    /// Starts a source-bound transactional edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the source section cannot be cloned for isolation.
    pub fn transaction(&self) -> Result<super::transaction::Transaction<'_>, OleError> {
        super::transaction::Transaction::from_snapshot(self)
    }

    /// Consumes the view into its complete generic section.
    #[must_use]
    pub fn into_section(self) -> Section {
        self.section
    }
}
