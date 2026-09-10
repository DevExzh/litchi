//! Typed parsing of the compressed MS-OVBA `dir` stream.

use super::{Error, Limits, check_limit, codec, invalid};
use litchi_codepage::Mbcs;

const PROJECT_CODEPAGE_ID: u16 = 0x0003;
const PROJECT_NAME_ID: u16 = 0x0004;
const PROJECT_SYS_KIND_ID: u16 = 0x0001;
const PROJECT_COMPAT_VERSION_ID: u16 = 0x004a;
const PROJECT_LCID_ID: u16 = 0x0002;
const PROJECT_LCID_INVOKE_ID: u16 = 0x0014;
const PROJECT_DOC_STRING_ID: u16 = 0x0005;
const PROJECT_HELP_FILE_PATH_ID: u16 = 0x0006;
const PROJECT_HELP_CONTEXT_ID: u16 = 0x0007;
const PROJECT_LIB_FLAGS_ID: u16 = 0x0008;
const PROJECT_VERSION_ID: u16 = 0x0009;
const PROJECT_CONSTANTS_ID: u16 = 0x000c;
const PROJECT_MODULES_ID: u16 = 0x000f;
const PROJECT_COOKIE_ID: u16 = 0x0013;
const MODULE_NAME_ID: u16 = 0x0019;
const MODULE_STREAM_NAME_ID: u16 = 0x001a;
const MODULE_DOC_STRING_ID: u16 = 0x001c;
const MODULE_HELP_CONTEXT_ID: u16 = 0x001e;
const MODULE_PROCEDURAL_ID: u16 = 0x0021;
const MODULE_OTHER_ID: u16 = 0x0022;
const MODULE_READ_ONLY_ID: u16 = 0x0025;
const MODULE_PRIVATE_ID: u16 = 0x0028;
const MODULE_TERMINATOR_ID: u16 = 0x002b;
const MODULE_COOKIE_ID: u16 = 0x002c;
const MODULE_OFFSET_ID: u16 = 0x0031;
const MODULE_NAME_UNICODE_ID: u16 = 0x0047;
const DIR_TERMINATOR_ID: u16 = 0x0010;
const REFERENCE_NAME_ID: u16 = 0x0016;
const REFERENCE_CONTROL_ID: u16 = 0x002f;
const REFERENCE_ORIGINAL_ID: u16 = 0x0033;
const REFERENCE_REGISTERED_ID: u16 = 0x000d;
const REFERENCE_PROJECT_ID: u16 = 0x000e;
const REFERENCE_NAME_RESERVED: u16 = 0x003e;
const REFERENCE_CONTROL_RESERVED: u16 = 0x0030;

const STREAM_NAME_RESERVED: u16 = 0x0032;
const PROJECT_DOC_STRING_RESERVED: u16 = 0x0040;
const PROJECT_HELP_FILE_PATH_RESERVED: u16 = 0x003d;
const PROJECT_CONSTANTS_RESERVED: u16 = 0x003c;
const MODULE_DOC_STRING_RESERVED: u16 = 0x0048;
const FIXED_U32_SIZE: u32 = 4;
const FIXED_U16_SIZE: u32 = 2;
const DEFAULT_PROJECT_LCID: u32 = 0x0409;
const WRITE_COOKIE: u16 = 0xffff;
const MAX_MODULE_IDENTIFIER_CHARACTERS: usize = 31;
const MAX_CFB_NAME_CODE_UNITS: usize = 31;
// `Mbcs::encode` returns an owned buffer for non-ASCII input.  Keep the
// preflight chunks small so a rejected aggregate limit never materializes the
// complete source string through the code-page adapter.  All byte-stream
// codecs accepted by `Mbcs` are stateless, so measuring at UTF-8 character
// boundaries preserves the exact strict encoder semantics.
const MBCS_PREFLIGHT_CHUNK_BYTES: usize = 256;

pub(crate) struct WriteProject<'a> {
    pub(crate) system_kind: u32,
    pub(crate) page: Mbcs,
    pub(crate) name: &'a str,
    pub(crate) description: &'a str,
    pub(crate) help_context: u32,
    pub(crate) version_major: u32,
    pub(crate) version_minor: u16,
    pub(crate) references: &'a [Reference],
    pub(crate) modules: &'a [WriteModule<'a>],
}

pub(crate) struct WriteModule<'a> {
    pub(crate) name: &'a str,
    pub(crate) stream_name: &'a str,
    pub(crate) description: &'a str,
    pub(crate) help_context: u32,
    pub(crate) kind: Kind,
    pub(crate) read_only: bool,
    pub(crate) private: bool,
}

/// One external reference declared by a VBA project.
///
/// References are metadata only.  litchi-vba never resolves, loads, or
/// instantiates a referenced type library, workbook, project, or control.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    name: Option<String>,
    kind: ReferenceKind,
    wire: Option<Vec<u8>>,
}

impl Reference {
    /// Construct a registered Automation type-library reference.
    pub fn registered(name: impl Into<String>, libid: impl Into<String>) -> Self {
        Self {
            name: Some(name.into()),
            kind: ReferenceKind::Registered {
                libid: libid.into(),
            },
            wire: None,
        }
    }

    /// Construct a reference to another VBA project.
    pub fn project(
        name: impl Into<String>,
        libid_absolute: impl Into<String>,
        libid_relative: impl Into<String>,
        major_version: u32,
        minor_version: u16,
    ) -> Self {
        Self {
            name: Some(name.into()),
            kind: ReferenceKind::Project {
                libid_absolute: libid_absolute.into(),
                libid_relative: libid_relative.into(),
                major_version,
                minor_version,
            },
            wire: None,
        }
    }

    /// Construct a twiddled/extended type-library reference.
    pub fn control(
        name: impl Into<String>,
        libid_twiddled: impl Into<String>,
        extended: Option<ExtendedReference>,
    ) -> Self {
        Self {
            name: Some(name.into()),
            kind: ReferenceKind::Control {
                libid_twiddled: libid_twiddled.into(),
                extended,
            },
            wire: None,
        }
    }

    /// Construct a reference-original record and its nested control record.
    pub fn original(
        name: Option<String>,
        libid_original: impl Into<String>,
        control: ControlReference,
    ) -> Self {
        Self {
            name,
            kind: ReferenceKind::Original {
                libid_original: libid_original.into(),
                control,
            },
            wire: None,
        }
    }

    /// Construct a bounded opaque reference record for an extension id.
    ///
    /// `payload` is the bytes following the record id and size fields.  This
    /// is intended for passive round-tripping of a producer extension whose
    /// semantics are outside this crate.
    pub fn opaque(id: u16, payload: impl Into<Vec<u8>>) -> Self {
        Self {
            name: None,
            kind: ReferenceKind::Opaque {
                id,
                payload: payload.into(),
            },
            wire: None,
        }
    }

    /// Optional human-readable reference name from `REFERENCENAME`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Typed reference record kind.
    #[must_use]
    pub fn kind(&self) -> &ReferenceKind {
        &self.kind
    }

    /// Exact source bytes for a reference read from a `dir` stream.
    ///
    /// Builders created with the constructors above return `None`.  Keeping
    /// this wire span lets a later source-bound edit retain producer-specific
    /// reserved values and record ordering when the reference is unchanged.
    #[must_use]
    pub fn raw(&self) -> Option<&[u8]> {
        self.wire.as_deref()
    }

    /// Return a copy with a replacement display name, clearing source bytes
    /// because the record must be re-encoded.
    #[must_use]
    pub fn with_name(mut self, name: Option<impl Into<String>>) -> Self {
        let name = name.map(Into::into);
        if self.name == name {
            return self;
        }
        self.name = name;
        self.wire = None;
        self
    }

    /// Replace the display name while retaining the source reference record's
    /// reserved fields and all following producer bytes.
    pub fn edit_name_source_bound(
        mut self,
        name: Option<impl Into<String>>,
        encoding: Mbcs,
        limits: &Limits,
    ) -> Result<Self, Error> {
        let name = name.map(Into::into);
        if self.name == name {
            return Ok(self);
        }
        if let Some(name) = name.as_deref() {
            validate_source_bound_reference_name(name, &self.kind)?;
        }
        let Some(wire) = self.wire.as_deref() else {
            self.name = name;
            return Ok(self);
        };
        let Some((name_end, name_reserved)) = reference_name_end(wire, encoding, limits)? else {
            if name.is_some() {
                let mut prefix = Vec::new();
                encode_reference_name(
                    &mut prefix,
                    name.as_deref(),
                    REFERENCE_NAME_RESERVED,
                    encoding,
                    limits,
                    source_bound_reference_name_kind(&self.kind),
                )?;
                prefix.extend_from_slice(wire);
                self.wire = Some(prefix);
            }
            self.name = name;
            return Ok(self);
        };
        let mut updated = Vec::new();
        encode_reference_name(
            &mut updated,
            name.as_deref(),
            name_reserved,
            encoding,
            limits,
            source_bound_reference_name_kind(&self.kind),
        )?;
        updated.extend_from_slice(&wire[name_end..]);
        self.name = name;
        self.wire = Some(updated);
        Ok(self)
    }
}

/// Typed record carried by [`Reference`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ReferenceKind {
    /// `REFERENCEREGISTERED` (0x000D).
    Registered { libid: String },
    /// `REFERENCEPROJECT` (0x000E).
    Project {
        libid_absolute: String,
        libid_relative: String,
        major_version: u32,
        minor_version: u16,
    },
    /// `REFERENCECONTROL` (0x002F).
    Control {
        libid_twiddled: String,
        extended: Option<ExtendedReference>,
    },
    /// `REFERENCEORIGINAL` (0x0033), with its nested control record.
    Original {
        libid_original: String,
        control: ControlReference,
    },
    /// A bounded record id not understood by this version of litchi-vba.
    Opaque { id: u16, payload: Vec<u8> },
}

/// Extended type-library information nested in a `REFERENCECONTROL` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtendedReference {
    name: Option<String>,
    libid: String,
    original_type_library: [u8; 16],
    cookie: u32,
}

impl ExtendedReference {
    /// Construct extended type-library metadata.
    pub fn new(
        name: Option<String>,
        libid: impl Into<String>,
        original_type_library: [u8; 16],
        cookie: u32,
    ) -> Self {
        Self {
            name,
            libid: libid.into(),
            original_type_library,
            cookie,
        }
    }

    /// Optional name of the extended type library.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Extended type-library identifier.
    #[must_use]
    pub fn libid(&self) -> &str {
        &self.libid
    }

    /// GUID of the original Automation type library.
    #[must_use]
    pub const fn original_type_library(&self) -> [u8; 16] {
        self.original_type_library
    }

    /// Extended type-library cookie.
    #[must_use]
    pub const fn cookie(&self) -> u32 {
        self.cookie
    }
}

/// A nested `REFERENCECONTROL` carried by `REFERENCEORIGINAL`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ControlReference {
    libid_twiddled: String,
    extended: Option<ExtendedReference>,
}

impl ControlReference {
    /// Construct a nested control record.
    pub fn new(libid_twiddled: impl Into<String>, extended: Option<ExtendedReference>) -> Self {
        Self {
            libid_twiddled: libid_twiddled.into(),
            extended,
        }
    }

    /// Twiddled type-library identifier.
    #[must_use]
    pub fn libid_twiddled(&self) -> &str {
        &self.libid_twiddled
    }

    /// Optional extended type-library metadata.
    #[must_use]
    pub fn extended(&self) -> Option<&ExtendedReference> {
        self.extended.as_ref()
    }
}

/// One mapping from the MBCS module name to its UTF-16 spelling in `PROJECTwm`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NameMap {
    mbcs_name: String,
    unicode_name: String,
}

impl NameMap {
    /// Construct a module-name map for a `PROJECTwm` stream.
    #[must_use]
    pub fn new(mbcs_name: impl Into<String>, unicode_name: impl Into<String>) -> Self {
        Self {
            mbcs_name: mbcs_name.into(),
            unicode_name: unicode_name.into(),
        }
    }

    /// Module name decoded from the project code page.
    #[must_use]
    pub fn mbcs_name(&self) -> &str {
        &self.mbcs_name
    }

    /// Module name decoded from UTF-16.
    #[must_use]
    pub fn unicode_name(&self) -> &str {
        &self.unicode_name
    }
}

/// Parsed metadata from an MS-OVBA `dir` stream, including external references
/// and module declarations.
#[derive(Debug, PartialEq, Eq)]
pub struct Dir {
    page: Mbcs,
    project_name: String,
    references: Vec<Reference>,
    modules: Vec<Module>,
}

impl Dir {
    /// Parse a complete compressed `dir` stream.
    ///
    /// # Errors
    ///
    /// Returns an error if the compressed container or the decompressed
    /// records are malformed, a configured [`Limits`] ceiling is exceeded, or
    /// the declared code page is unsupported.
    pub fn parse(compressed: &[u8], limits: &Limits) -> Result<Self, Error> {
        let decompressed = codec::decode(compressed, limits)?;
        Self::parse_decompressed(&decompressed, limits)
    }

    /// Checked page used by MBCS strings and module source.
    #[must_use]
    pub fn page(&self) -> Mbcs {
        self.page
    }

    /// Numeric page identifier stored in `PROJECTCODEPAGE`.
    #[must_use]
    pub fn page_id(&self) -> u16 {
        self.page.id16()
    }

    /// VBA project identifier from `PROJECTNAME`.
    #[must_use]
    pub fn project_name(&self) -> &str {
        &self.project_name
    }

    /// Module metadata in directory order.
    #[must_use]
    pub fn modules(&self) -> &[Module] {
        &self.modules
    }

    /// External Automation type-library and VBA-project references in order.
    #[must_use]
    pub fn references(&self) -> &[Reference] {
        &self.references
    }

    pub(crate) fn parse_decompressed(data: &[u8], limits: &Limits) -> Result<Self, Error> {
        Self::parse_decompressed_with_end_policy(data, limits, false)
    }

    pub(crate) fn parse_decompressed_strict(data: &[u8], limits: &Limits) -> Result<Self, Error> {
        Self::parse_decompressed_with_end_policy(data, limits, true)
    }

    fn parse_decompressed_with_end_policy(
        data: &[u8],
        limits: &Limits,
        require_end: bool,
    ) -> Result<Self, Error> {
        let (information_end, page, project_name_bytes) = parse_project_information(data, limits)?;
        let project_name = decode_mbcs(&project_name_bytes, page, "PROJECTNAME")?;
        let (references, modules) =
            find_references_and_modules(&data[information_end..], page, limits, require_end)?;
        Ok(Self {
            page,
            project_name,
            references,
            modules,
        })
    }
}

/// Broad module category encoded by `MODULETYPE`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    /// Standard procedural module.
    Procedural,
    /// Document, class, or designer module.
    DocumentClassOrDesigner,
}

/// Metadata locating one module's inert source stream.
#[derive(Debug, PartialEq, Eq)]
pub struct Module {
    name: String,
    stream_name: String,
    text_offset: u32,
    kind: Kind,
    read_only: bool,
    private: bool,
}

impl Module {
    /// VBA identifier for this module.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// CFB stream name containing this module.
    #[must_use]
    pub fn stream_name(&self) -> &str {
        &self.stream_name
    }

    /// Byte offset at which compressed source begins in the module stream.
    #[must_use]
    pub fn text_offset(&self) -> u32 {
        self.text_offset
    }

    /// Broad module category from `MODULETYPE`.
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Whether `MODULEREADONLY` is present.
    #[must_use]
    pub fn is_read_only(&self) -> bool {
        self.read_only
    }

    /// Whether `MODULEPRIVATE` is present.
    #[must_use]
    pub fn is_private(&self) -> bool {
        self.private
    }
}

struct Reader<'a> {
    data: &'a [u8],
    position: usize,
}

impl<'a> Reader<'a> {
    fn new(data: &'a [u8], position: usize) -> Self {
        Self { data, position }
    }

    fn peek_u16(&self) -> Option<u16> {
        read_u16_at(self.data, self.position)
    }

    fn read_u16(&mut self) -> Result<u16, Error> {
        let value = read_u16_at(self.data, self.position)
            .ok_or_else(|| invalid("truncated 16-bit dir-stream field"))?;
        self.position += 2;
        Ok(value)
    }

    fn read_u32(&mut self) -> Result<u32, Error> {
        let value = read_u32_at(self.data, self.position)
            .ok_or_else(|| invalid("truncated 32-bit dir-stream field"))?;
        self.position += 4;
        Ok(value)
    }

    fn read_bytes(&mut self, length: usize) -> Result<&'a [u8], Error> {
        let end = self
            .position
            .checked_add(length)
            .ok_or_else(|| invalid("dir-stream field length overflow"))?;
        let value = self
            .data
            .get(self.position..end)
            .ok_or_else(|| invalid("truncated dir-stream field"))?;
        self.position = end;
        Ok(value)
    }

    fn expect_id(&mut self, expected: u16) -> Result<(), Error> {
        let actual = self.read_u16()?;
        if actual != expected {
            return Err(invalid(format!(
                "expected dir record {expected:#06x}, found {actual:#06x}"
            )));
        }
        Ok(())
    }

    fn expect_u32(&mut self, expected: u32, field: &'static str) -> Result<(), Error> {
        let actual = self.read_u32()?;
        if actual != expected {
            return Err(invalid(format!(
                "{field} must be {expected:#010x}, found {actual:#010x}"
            )));
        }
        Ok(())
    }

    fn expect_sized_u32(&mut self, id: u16, size: u32) -> Result<(), Error> {
        self.expect_id(id)?;
        self.expect_u32(size, "dir record size")
    }

    fn length_prefixed(&mut self, limits: &Limits) -> Result<&'a [u8], Error> {
        let Ok(length) = usize::try_from(self.read_u32()?) else {
            return Err(invalid("dir-stream string length does not fit usize"));
        };
        check_limit("VBA string bytes", length, limits.max_string_bytes)?;
        self.read_bytes(length)
    }

    fn length_prefixed_bounded(
        &mut self,
        limits: &Limits,
        protocol_maximum: usize,
        field: &'static str,
    ) -> Result<&'a [u8], Error> {
        let value = self.length_prefixed(limits)?;
        if value.len() > protocol_maximum {
            return Err(invalid(format!(
                "{field} length {} exceeds {protocol_maximum}",
                value.len()
            )));
        }
        Ok(value)
    }

    fn string_pair(
        &mut self,
        encoding: Mbcs,
        reserved: u16,
        field: &'static str,
        limits: &Limits,
    ) -> Result<String, Error> {
        self.string_pair_bounded(encoding, reserved, field, limits, usize::MAX)
    }

    fn string_pair_bounded(
        &mut self,
        encoding: Mbcs,
        _reserved: u16,
        field: &'static str,
        limits: &Limits,
        protocol_maximum: usize,
    ) -> Result<String, Error> {
        let mbcs = decode_mbcs(
            self.length_prefixed_bounded(limits, protocol_maximum, field)?,
            encoding,
            field,
        )?;
        // Reserved words are canonical on write but explicitly ignored on
        // read by MS-OVBA. Consuming them without validation keeps
        // source-backed replay lossless.
        let _actual_reserved = self.read_u16()?;
        let unicode_bytes =
            self.length_prefixed_bounded(limits, protocol_maximum.saturating_mul(2), field)?;
        let unicode = decode_utf16(unicode_bytes, field)?;
        if unicode != mbcs {
            return Err(invalid(format!(
                "{field} Unicode value does not match its MBCS value"
            )));
        }
        Ok(unicode)
    }

    fn mbcs_pair_bounded(
        &mut self,
        encoding: Mbcs,
        _reserved: u16,
        field: &'static str,
        limits: &Limits,
        protocol_maximum: usize,
    ) -> Result<String, Error> {
        let first_bytes = self.length_prefixed_bounded(limits, protocol_maximum, field)?;
        let first = decode_mbcs(first_bytes, encoding, field)?;
        let _actual_reserved = self.read_u16()?;
        let second_bytes = self.length_prefixed_bounded(limits, protocol_maximum, field)?;
        let second = decode_mbcs(second_bytes, encoding, field)?;
        if first_bytes != second_bytes || first != second {
            return Err(invalid(format!(
                "{field} duplicate path values do not match"
            )));
        }
        Ok(first)
    }
}

pub(crate) fn encode_dir(project: &WriteProject<'_>, limits: &Limits) -> Result<Vec<u8>, Error> {
    check_limit(
        "VBA module count",
        project.modules.len(),
        limits.max_modules,
    )?;
    check_limit(
        "VBA reference count",
        project.references.len(),
        limits.max_references,
    )?;
    let Ok(module_count) = u16::try_from(project.modules.len()) else {
        return Err(invalid("VBA module count exceeds the dir-stream field"));
    };
    let encoding = project.page;
    let mut output = Vec::new();

    push_record_bounded(
        &mut output,
        PROJECT_SYS_KIND_ID,
        &project.system_kind.to_le_bytes(),
        limits,
    )?;
    push_record_bounded(
        &mut output,
        PROJECT_LCID_ID,
        &DEFAULT_PROJECT_LCID.to_le_bytes(),
        limits,
    )?;
    push_record_bounded(
        &mut output,
        PROJECT_LCID_INVOKE_ID,
        &DEFAULT_PROJECT_LCID.to_le_bytes(),
        limits,
    )?;
    push_record_bounded(
        &mut output,
        PROJECT_CODEPAGE_ID,
        &project.page.id16().to_le_bytes(),
        limits,
    )?;

    validate_vba_identifier(project.name, "PROJECTNAME")?;
    let project_name = encode_mbcs(project.name, encoding, "PROJECTNAME")?;
    check_protocol_length("PROJECTNAME", project_name.len(), 128)?;
    push_record_bounded(&mut output, PROJECT_NAME_ID, &project_name, limits)?;
    push_string_pair(
        &mut output,
        PROJECT_DOC_STRING_ID,
        PROJECT_DOC_STRING_RESERVED,
        project.description,
        encoding,
        limits,
        2_000,
        "PROJECTDOCSTRING",
    )?;
    push_mbcs_pair(
        &mut output,
        PROJECT_HELP_FILE_PATH_ID,
        PROJECT_HELP_FILE_PATH_RESERVED,
        "",
        encoding,
        limits,
        260,
        "PROJECTHELPFILEPATH",
    )?;
    push_record_bounded(
        &mut output,
        PROJECT_HELP_CONTEXT_ID,
        &project.help_context.to_le_bytes(),
        limits,
    )?;
    push_record_bounded(
        &mut output,
        PROJECT_LIB_FLAGS_ID,
        &0u32.to_le_bytes(),
        limits,
    )?;
    append_bounded(&mut output, &PROJECT_VERSION_ID.to_le_bytes(), limits)?;
    append_bounded(&mut output, &FIXED_U32_SIZE.to_le_bytes(), limits)?;
    append_bounded(&mut output, &project.version_major.to_le_bytes(), limits)?;
    append_bounded(&mut output, &project.version_minor.to_le_bytes(), limits)?;

    for reference in project.references {
        encode_reference(&mut output, reference, encoding, limits)?;
        check_limit(
            "decompressed VBA stream bytes",
            output.len(),
            limits.max_decompressed_stream_bytes,
        )?;
    }

    push_record_bounded(
        &mut output,
        PROJECT_MODULES_ID,
        &module_count.to_le_bytes(),
        limits,
    )?;
    push_record_bounded(
        &mut output,
        PROJECT_COOKIE_ID,
        &WRITE_COOKIE.to_le_bytes(),
        limits,
    )?;
    for module in project.modules {
        encode_module(&mut output, module, encoding, limits)?;
        check_limit(
            "decompressed VBA stream bytes",
            output.len(),
            limits.max_decompressed_stream_bytes,
        )?;
    }
    append_bounded(&mut output, &DIR_TERMINATOR_ID.to_le_bytes(), limits)?;
    append_bounded(&mut output, &0u32.to_le_bytes(), limits)?;

    check_limit(
        "decompressed VBA stream bytes",
        output.len(),
        limits.max_decompressed_stream_bytes,
    )?;
    codec::encode(&output, limits)
}

/// Re-encode only the source-bound reference array in an existing `dir`
/// stream.  The project-information and module records remain byte-for-byte
/// unchanged, including producer-specific reserved values and records this
/// crate does not model.
pub(crate) fn encode_dir_with_references(
    source: &[u8],
    references: &[Reference],
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    check_limit(
        "decompressed VBA stream bytes",
        source.len(),
        limits.max_decompressed_stream_bytes,
    )?;
    check_limit(
        "VBA reference count",
        references.len(),
        limits.max_references,
    )?;
    let (information_end, encoding, _) = parse_project_information(source, limits)?;
    let mut reader = Reader::new(source, information_end);
    let references_start = reader.position;
    let mut source_reference_count = 0usize;
    loop {
        let id = reader
            .peek_u16()
            .ok_or_else(|| invalid("missing PROJECTMODULES record"))?;
        if id == PROJECT_MODULES_ID {
            break;
        }
        source_reference_count = source_reference_count
            .checked_add(1)
            .ok_or_else(|| invalid("VBA reference count overflow"))?;
        check_limit(
            "VBA reference count",
            source_reference_count,
            limits.max_references,
        )?;
        parse_reference(&mut reader, encoding, limits)?;
    }
    let references_end = reader.position;

    // A changed-reference edit preserves the source module suffix, so charge
    // its declared module count under this operation's limits before any
    // output allocation.  The no-op path receives the same check through
    // `Dir::parse_decompressed`.
    let mut module_reader = Reader::new(source, references_end);
    module_reader.expect_sized_u32(PROJECT_MODULES_ID, FIXED_U16_SIZE)?;
    let source_module_count = usize::from(module_reader.read_u16()?);
    check_limit("VBA module count", source_module_count, limits.max_modules)?;

    let base_len = source
        .len()
        .checked_sub(references_end.saturating_sub(references_start))
        .ok_or_else(|| invalid("dir reference span exceeds source length"))?;
    let reference_limits = Limits {
        max_decompressed_stream_bytes: limits
            .max_decompressed_stream_bytes
            .saturating_sub(base_len),
        ..*limits
    };
    let mut encoded_references = Vec::new();
    for reference in references {
        encode_reference(
            &mut encoded_references,
            reference,
            encoding,
            &reference_limits,
        )?;
    }
    let new_len = source
        .len()
        .checked_sub(references_end.saturating_sub(references_start))
        .and_then(|length| length.checked_add(encoded_references.len()))
        .ok_or_else(|| invalid("dir stream size overflow"))?;
    check_limit(
        "decompressed VBA stream bytes",
        new_len,
        limits.max_decompressed_stream_bytes,
    )?;
    let mut updated = Vec::with_capacity(new_len);
    updated.extend_from_slice(&source[..references_start]);
    updated.extend_from_slice(&encoded_references);
    updated.extend_from_slice(&source[references_end..]);
    codec::encode(&updated, limits)
}

fn encode_module(
    output: &mut Vec<u8>,
    module: &WriteModule<'_>,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<(), Error> {
    validate_module_identifier(module.name, "MODULENAME")?;
    let name = encode_mbcs(module.name, encoding, "MODULENAME")?;
    check_limit("VBA string bytes", name.len(), limits.max_string_bytes)?;
    push_record_bounded(output, MODULE_NAME_ID, &name, limits)?;
    let unicode_name = encode_utf16(module.name, "MODULENAMEUNICODE")?;
    check_limit(
        "VBA string bytes",
        unicode_name.len(),
        limits.max_string_bytes,
    )?;
    push_record_bounded(output, MODULE_NAME_UNICODE_ID, &unicode_name, limits)?;
    push_string_pair(
        output,
        MODULE_STREAM_NAME_ID,
        STREAM_NAME_RESERVED,
        module.stream_name,
        encoding,
        limits,
        limits.max_string_bytes,
        "MODULESTREAMNAME",
    )?;
    push_string_pair(
        output,
        MODULE_DOC_STRING_ID,
        MODULE_DOC_STRING_RESERVED,
        module.description,
        encoding,
        limits,
        limits.max_string_bytes,
        "MODULEDOCSTRING",
    )?;
    push_record_bounded(output, MODULE_OFFSET_ID, &0u32.to_le_bytes(), limits)?;
    push_record_bounded(
        output,
        MODULE_HELP_CONTEXT_ID,
        &module.help_context.to_le_bytes(),
        limits,
    )?;
    push_record_bounded(
        output,
        MODULE_COOKIE_ID,
        &WRITE_COOKIE.to_le_bytes(),
        limits,
    )?;
    let type_id = match module.kind {
        Kind::Procedural => MODULE_PROCEDURAL_ID,
        Kind::DocumentClassOrDesigner => MODULE_OTHER_ID,
    };
    append_bounded(output, &type_id.to_le_bytes(), limits)?;
    append_bounded(output, &0u32.to_le_bytes(), limits)?;
    if module.read_only {
        append_bounded(output, &MODULE_READ_ONLY_ID.to_le_bytes(), limits)?;
        append_bounded(output, &0u32.to_le_bytes(), limits)?;
    }
    if module.private {
        append_bounded(output, &MODULE_PRIVATE_ID.to_le_bytes(), limits)?;
        append_bounded(output, &0u32.to_le_bytes(), limits)?;
    }
    append_bounded(output, &MODULE_TERMINATOR_ID.to_le_bytes(), limits)?;
    append_bounded(output, &0u32.to_le_bytes(), limits)?;
    Ok(())
}

fn encode_reference(
    output: &mut Vec<u8>,
    reference: &Reference,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<(), Error> {
    if let Some(wire) = reference.wire.as_deref() {
        let new_len = output
            .len()
            .checked_add(wire.len())
            .ok_or_else(|| invalid("VBA reference stream size overflow"))?;
        check_limit(
            "decompressed VBA stream bytes",
            new_len,
            limits.max_decompressed_stream_bytes,
        )?;
        output.extend_from_slice(wire);
        return Ok(());
    }
    let encoded_size = encoded_reference_size(reference, encoding, limits)?;
    let new_len = output
        .len()
        .checked_add(encoded_size)
        .ok_or_else(|| invalid("VBA reference stream size overflow"))?;
    check_limit(
        "decompressed VBA stream bytes",
        new_len,
        limits.max_decompressed_stream_bytes,
    )?;
    let mut encoded = Vec::new();
    encode_reference_inner(&mut encoded, reference, encoding, limits)?;
    output.extend_from_slice(&encoded);
    Ok(())
}

fn encoded_reference_size(
    reference: &Reference,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<usize, Error> {
    let name_size = reference
        .name
        .as_deref()
        .map(|name| {
            validate_reference_identifier(
                name,
                reference_name_kind(&reference.kind),
                "REFERENCENAME",
            )?;
            encoded_string_pair_size(
                name,
                encoding,
                limits,
                limits.max_string_bytes,
                "REFERENCENAME",
            )
        })
        .transpose()?
        .unwrap_or(0);
    let body_size = match &reference.kind {
        ReferenceKind::Registered { libid } => {
            let libid_len =
                encoded_libid_reference_len(libid, encoding, limits, "REFERENCEREGISTERED Libid")?;
            16usize
                .checked_add(libid_len)
                .ok_or_else(|| invalid("VBA registered-reference size overflow"))?
        },
        ReferenceKind::Project {
            libid_absolute,
            libid_relative,
            ..
        } => {
            let absolute_len = encoded_project_reference_len(
                libid_absolute,
                encoding,
                limits,
                "REFERENCEPROJECT LibidAbsolute",
            )?;
            let relative_len = encoded_project_reference_len(
                libid_relative,
                encoding,
                limits,
                "REFERENCEPROJECT LibidRelative",
            )?;
            20usize
                .checked_add(absolute_len)
                .and_then(|value| value.checked_add(relative_len))
                .ok_or_else(|| invalid("VBA project-reference size overflow"))?
        },
        ReferenceKind::Control {
            libid_twiddled,
            extended,
        } => encoded_control_reference_size(libid_twiddled, extended.as_ref(), encoding, limits)?,
        ReferenceKind::Original {
            libid_original,
            control,
        } => {
            let original_len = encoded_libid_reference_len(
                libid_original,
                encoding,
                limits,
                "REFERENCEORIGINAL LibidOriginal",
            )?;
            let control_size = encoded_control_reference_size(
                &control.libid_twiddled,
                control.extended.as_ref(),
                encoding,
                limits,
            )?;
            6usize
                .checked_add(original_len)
                .and_then(|value| value.checked_add(control_size))
                .ok_or_else(|| invalid("VBA original-reference size overflow"))?
        },
        ReferenceKind::Opaque { payload, .. } => 6usize
            .checked_add(payload.len())
            .ok_or_else(|| invalid("VBA opaque-reference size overflow"))?,
    };
    name_size
        .checked_add(body_size)
        .ok_or_else(|| invalid("VBA reference stream size overflow"))
}

fn encoded_control_reference_size(
    libid_twiddled: &str,
    extended: Option<&ExtendedReference>,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<usize, Error> {
    let twiddled_len = encoded_libid_reference_len(
        libid_twiddled,
        encoding,
        limits,
        "REFERENCECONTROL LibidTwiddled",
    )?;
    let extended = extended
        .ok_or_else(|| invalid("REFERENCECONTROL requires extended type-library metadata"))?;
    let extended_libid_len = encoded_libid_reference_len(
        &extended.libid,
        encoding,
        limits,
        "REFERENCECONTROL LibidExtended",
    )?;
    let extended_name = extended
        .name
        .as_deref()
        .map(|name| {
            validate_reference_identifier(name, ReferenceNameKind::Library, "REFERENCENAME")?;
            encoded_string_pair_size(
                name,
                encoding,
                limits,
                limits.max_string_bytes,
                "REFERENCENAME",
            )
        })
        .transpose()?
        .unwrap_or(0);
    52usize
        .checked_add(twiddled_len)
        .and_then(|value| value.checked_add(extended_libid_len))
        .and_then(|value| value.checked_add(extended_name))
        .ok_or_else(|| invalid("VBA control-reference size overflow"))
}

fn encoded_string_pair_size(
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    protocol_maximum: usize,
    field: &'static str,
) -> Result<usize, Error> {
    check_limit(
        "VBA input string bytes",
        value.len(),
        limits.max_string_bytes.saturating_mul(4),
    )?;
    let mbcs_len = encoded_reference_value_len(value, encoding, limits, field, |_| Ok(()))?;
    check_protocol_length(field, mbcs_len, protocol_maximum)?;
    let unicode_len = value
        .encode_utf16()
        .count()
        .checked_mul(2)
        .ok_or_else(|| invalid(format!("{field} UTF-16 length overflow")))?;
    check_limit("VBA string bytes", unicode_len, limits.max_string_bytes)?;
    check_protocol_length(field, unicode_len, protocol_maximum.saturating_mul(2))?;
    12usize
        .checked_add(mbcs_len)
        .and_then(|value| value.checked_add(unicode_len))
        .ok_or_else(|| invalid(format!("{field} encoded size overflow")))
}

fn reference_name_end(
    wire: &[u8],
    encoding: Mbcs,
    limits: &Limits,
) -> Result<Option<(usize, u16)>, Error> {
    if wire.len() < 2 || read_u16_at(wire, 0) != Some(REFERENCE_NAME_ID) {
        return Ok(None);
    }
    let mut reader = Reader::new(wire, 0);
    reader.expect_id(REFERENCE_NAME_ID)?;
    let mbcs = decode_mbcs(reader.length_prefixed(limits)?, encoding, "REFERENCENAME")?;
    let reserved = reader.read_u16()?;
    let unicode = decode_utf16(reader.length_prefixed(limits)?, "REFERENCENAMEUNICODE")?;
    if unicode != mbcs {
        return Err(invalid("REFERENCENAMEUNICODE does not match REFERENCENAME"));
    }
    Ok(Some((reader.position, reserved)))
}

fn encode_reference_name(
    output: &mut Vec<u8>,
    name: Option<&str>,
    reserved: u16,
    encoding: Mbcs,
    limits: &Limits,
    kind: ReferenceNameKind,
) -> Result<(), Error> {
    if let Some(name) = name {
        validate_reference_identifier(name, kind, "REFERENCENAME")?;
        push_string_pair(
            output,
            REFERENCE_NAME_ID,
            reserved,
            name,
            encoding,
            limits,
            limits.max_string_bytes,
            "REFERENCENAME",
        )?;
    }
    Ok(())
}

fn encode_reference_inner(
    output: &mut Vec<u8>,
    reference: &Reference,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<(), Error> {
    let name_kind = reference_name_kind(&reference.kind);
    if let Some(name) = &reference.name {
        validate_reference_identifier(name, name_kind, "REFERENCENAME")?;
        push_string_pair(
            output,
            REFERENCE_NAME_ID,
            REFERENCE_NAME_RESERVED,
            name,
            encoding,
            limits,
            limits.max_string_bytes,
            "REFERENCENAME",
        )?;
    }
    match &reference.kind {
        ReferenceKind::Registered { libid } => {
            let encoded =
                encode_libid_reference_bytes(libid, encoding, limits, "REFERENCEREGISTERED Libid")?;
            let mut payload = Vec::new();
            push_length_prefixed(&mut payload, &encoded)?;
            payload.extend_from_slice(&0u32.to_le_bytes());
            payload.extend_from_slice(&0u16.to_le_bytes());
            push_record_bounded(output, REFERENCE_REGISTERED_ID, &payload, limits)?;
        },
        ReferenceKind::Project {
            libid_absolute,
            libid_relative,
            major_version,
            minor_version,
        } => {
            let absolute = encode_project_reference_bytes(
                libid_absolute,
                encoding,
                limits,
                "REFERENCEPROJECT LibidAbsolute",
            )?;
            let relative = encode_project_reference_bytes(
                libid_relative,
                encoding,
                limits,
                "REFERENCEPROJECT LibidRelative",
            )?;
            let mut payload = Vec::new();
            push_length_prefixed(&mut payload, &absolute)?;
            push_length_prefixed(&mut payload, &relative)?;
            payload.extend_from_slice(&major_version.to_le_bytes());
            payload.extend_from_slice(&minor_version.to_le_bytes());
            push_record_bounded(output, REFERENCE_PROJECT_ID, &payload, limits)?;
        },
        ReferenceKind::Control {
            libid_twiddled,
            extended,
        } => encode_control_reference(
            output,
            &ControlReference::new(libid_twiddled.clone(), extended.clone()),
            encoding,
            limits,
        )?,
        ReferenceKind::Original {
            libid_original,
            control,
        } => {
            let original = encode_libid_reference_bytes(
                libid_original,
                encoding,
                limits,
                "REFERENCEORIGINAL LibidOriginal",
            )?;
            append_bounded(output, &REFERENCE_ORIGINAL_ID.to_le_bytes(), limits)?;
            push_length_prefixed_bounded(output, &original, limits)?;
            encode_control_reference(output, control, encoding, limits)?;
        },
        ReferenceKind::Opaque { id, payload } => {
            if is_known_dir_record_id(*id) {
                return Err(invalid(format!(
                    "opaque VBA reference id {id:#06x} is a known directory record"
                )));
            }
            check_limit("VBA string bytes", payload.len(), limits.max_string_bytes)?;
            push_record_bounded(output, *id, payload, limits)?;
        },
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReferenceNameKind {
    Project,
    Library,
    SourceBoundLibrary,
}

fn reference_name_kind(kind: &ReferenceKind) -> ReferenceNameKind {
    match kind {
        ReferenceKind::Project { .. } => ReferenceNameKind::Project,
        ReferenceKind::Registered { .. }
        | ReferenceKind::Control { .. }
        | ReferenceKind::Original { .. }
        | ReferenceKind::Opaque { .. } => ReferenceNameKind::Library,
    }
}

fn source_bound_reference_name_kind(kind: &ReferenceKind) -> ReferenceNameKind {
    match kind {
        ReferenceKind::Project { .. } => ReferenceNameKind::Project,
        ReferenceKind::Registered { .. }
        | ReferenceKind::Control { .. }
        | ReferenceKind::Original { .. }
        | ReferenceKind::Opaque { .. } => ReferenceNameKind::SourceBoundLibrary,
    }
}

fn validate_source_bound_reference_name(name: &str, kind: &ReferenceKind) -> Result<(), Error> {
    validate_reference_identifier(
        name,
        source_bound_reference_name_kind(kind),
        "REFERENCENAME",
    )
}

fn validate_reference_identifier(
    name: &str,
    kind: ReferenceNameKind,
    field: &'static str,
) -> Result<(), Error> {
    let mut characters = name.chars();
    let Some(first) = characters.next() else {
        return Err(invalid(format!("{field} must not be empty")));
    };
    let valid_subsequent = match kind {
        ReferenceNameKind::Project => valid_identifier_subsequent,
        ReferenceNameKind::Library => valid_identifier_subsequent,
        // A source-bound edit may retain a producer's existing C706 display
        // spelling such as `stdole-edited`. This exception is unavailable to
        // fresh authoring paths.
        ReferenceNameKind::SourceBoundLibrary => {
            |character| valid_identifier_subsequent(character) || character == '-'
        },
    };
    if !valid_identifier_initial(first) || characters.any(|character| !valid_subsequent(character))
    {
        return Err(invalid(format!("{field} is not a reference identifier")));
    }
    if RESERVED_IDENTIFIERS
        .iter()
        .any(|reserved| reserved.eq_ignore_ascii_case(name))
    {
        return Err(invalid(format!("{field} is a reserved VBA identifier")));
    }
    Ok(())
}

pub(crate) fn validate_vba_identifier(value: &str, field: &'static str) -> Result<(), Error> {
    let mut characters = value.chars();
    let Some(first) = characters.next() else {
        return Err(invalid(format!("{field} must not be empty")));
    };
    if !valid_identifier_initial(first)
        || characters.any(|character| !valid_identifier_subsequent(character))
    {
        return Err(invalid(format!("{field} is not a VBA identifier")));
    }
    if RESERVED_IDENTIFIERS
        .iter()
        .any(|reserved| reserved.eq_ignore_ascii_case(value))
    {
        return Err(invalid(format!("{field} is a reserved VBA identifier")));
    }
    Ok(())
}

fn valid_identifier_initial(character: char) -> bool {
    character.is_ascii_alphabetic()
        || (!character.is_ascii() && !character.is_control() && !character.is_whitespace())
}

fn valid_identifier_subsequent(character: char) -> bool {
    valid_identifier_initial(character) || character.is_ascii_digit() || character == '_'
}

const RESERVED_IDENTIFIERS: &[&str] = &[
    "Call",
    "Case",
    "Close",
    "Const",
    "Declare",
    "DefBool",
    "DefByte",
    "DefCur",
    "DefDate",
    "DefDbl",
    "DefInt",
    "DefLng",
    "DefLngLng",
    "DefLngPtr",
    "DefObj",
    "DefSng",
    "DefStr",
    "DefVar",
    "Dim",
    "Do",
    "Else",
    "ElseIf",
    "End",
    "EndIf",
    "Enum",
    "Erase",
    "Event",
    "Exit",
    "For",
    "Friend",
    "Function",
    "Get",
    "Global",
    "GoSub",
    "GoTo",
    "If",
    "Implements",
    "Input",
    "Let",
    "Lock",
    "Loop",
    "LSet",
    "Next",
    "On",
    "Open",
    "Option",
    "Print",
    "Private",
    "Public",
    "Put",
    "RaiseEvent",
    "ReDim",
    "Resume",
    "Return",
    "RSet",
    "Seek",
    "Select",
    "Set",
    "Static",
    "Stop",
    "Sub",
    "Type",
    "Unlock",
    "Wend",
    "While",
    "With",
    "Write",
    "Rem",
    "Any",
    "As",
    "ByRef",
    "ByVal",
    "Each",
    "Else",
    "In",
    "New",
    "Shared",
    "Until",
    "WithEvents",
    "Optional",
    "ParamArray",
    "Preserve",
    "Spc",
    "Tab",
    "Then",
    "To",
    "AddressOf",
    "And",
    "Eqv",
    "Imp",
    "Is",
    "Like",
    "Mod",
    "Not",
    "Or",
    "TypeOf",
    "Xor",
    "Abs",
    "CBool",
    "CByte",
    "CCur",
    "CDate",
    "CDbl",
    "CDec",
    "CInt",
    "CLng",
    "CLngLng",
    "CLngPtr",
    "CSng",
    "CStr",
    "CVar",
    "CVErr",
    "Date",
    "Debug",
    "DoEvents",
    "Fix",
    "Int",
    "Len",
    "LenB",
    "Me",
    "PSet",
    "Scale",
    "Sgn",
    "String",
    "Array",
    "Circle",
    "InputB",
    "LBound",
    "UBound",
    "Boolean",
    "Byte",
    "Currency",
    "Double",
    "Integer",
    "Long",
    "LongLong",
    "LongPtr",
    "Single",
    "Variant",
    "True",
    "False",
    "Nothing",
    "Empty",
    "Null",
    "Attribute",
    "LINEINPUT",
    "VB_Base",
    "VB_Control",
    "VB_Creatable",
    "VB_Customizable",
    "VB_Description",
    "VB_Exposed",
    "VB_Ext_KEY",
    "VB_GlobalNameSpace",
    "VB_HelpID",
    "VB_Invoke_Func",
    "VB_MemberFlags",
    "VB_Name",
    "VB_PredeclaredId",
    "VB_ProcData",
    "VB_TemplateDerived",
    "VB_UserMemId",
    "VB_VarDescription",
    "VB_VarHelpID",
    "VB_VarMemberFlags",
    "VB_VarProcData",
    "VB_VarUserMemId",
    "CDecl",
    "Decimal",
    "DefDec",
];

fn is_known_dir_record_id(id: u16) -> bool {
    matches!(
        id,
        PROJECT_SYS_KIND_ID
            | PROJECT_COMPAT_VERSION_ID
            | PROJECT_LCID_ID
            | PROJECT_LCID_INVOKE_ID
            | PROJECT_CODEPAGE_ID
            | PROJECT_NAME_ID
            | PROJECT_DOC_STRING_ID
            | PROJECT_HELP_FILE_PATH_ID
            | PROJECT_HELP_CONTEXT_ID
            | PROJECT_LIB_FLAGS_ID
            | PROJECT_VERSION_ID
            | PROJECT_CONSTANTS_ID
            | PROJECT_MODULES_ID
            | PROJECT_COOKIE_ID
            | MODULE_NAME_ID
            | MODULE_STREAM_NAME_ID
            | MODULE_DOC_STRING_ID
            | MODULE_HELP_CONTEXT_ID
            | MODULE_PROCEDURAL_ID
            | MODULE_OTHER_ID
            | MODULE_READ_ONLY_ID
            | MODULE_PRIVATE_ID
            | MODULE_TERMINATOR_ID
            | MODULE_COOKIE_ID
            | MODULE_OFFSET_ID
            | MODULE_NAME_UNICODE_ID
            | DIR_TERMINATOR_ID
            | REFERENCE_NAME_ID
            | REFERENCE_CONTROL_ID
            | REFERENCE_ORIGINAL_ID
            | REFERENCE_REGISTERED_ID
            | REFERENCE_PROJECT_ID
    )
}

fn encode_reference_bytes(
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    field: &'static str,
) -> Result<Vec<u8>, Error> {
    check_limit(
        "VBA input string bytes",
        value.len(),
        limits.max_string_bytes.saturating_mul(4),
    )?;
    let bytes = encode_mbcs(value, encoding, field)?;
    check_limit("VBA string bytes", bytes.len(), limits.max_string_bytes)?;
    Ok(bytes)
}

/// Validate and measure an MBCS field without materializing the whole value.
/// Reference size precharging runs before the aggregate `dir` output check; it
/// must not materialize a large LibidReference only to have that projected
/// size check reject it.  Semantic Libid/Project grammar validation for a
/// non-ASCII value is repeated by the actual bounded writer after the
/// aggregate check; the chunked pass still rejects unmappable characters and
/// encoded NUL bytes before returning a size.
fn encoded_reference_value_len<F>(
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    field: &'static str,
    validate: F,
) -> Result<usize, Error>
where
    F: FnOnce(&[u8]) -> Result<(), Error>,
{
    let length = encoded_mbcs_len(value, encoding, limits, field)?;
    if value.is_ascii() {
        validate(value.as_bytes())?;
    }
    Ok(length)
}

/// Measure strict MBCS output in bounded chunks.
///
/// `Mbcs::encode` intentionally returns an owned `Cow` for non-ASCII input.
/// Calling it once for an attacker-sized value defeats the reference-size
/// precharge, even when the aggregate stream limit will reject the value.
/// Encoding at UTF-8 character boundaries keeps each temporary output bounded
/// while retaining the code-page adapter's unmappable-character behavior.
fn encoded_mbcs_len(
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    field: &'static str,
) -> Result<usize, Error> {
    check_limit(
        "VBA input string bytes",
        value.len(),
        limits.max_string_bytes.saturating_mul(4),
    )?;
    if value.contains('\0') {
        return Err(invalid(format!("{field} contains a null character")));
    }
    if value.is_ascii() {
        check_limit("VBA string bytes", value.len(), limits.max_string_bytes)?;
        return Ok(value.len());
    }

    let mut length = 0usize;
    let mut start = 0usize;
    while start < value.len() {
        let mut end = (start + MBCS_PREFLIGHT_CHUNK_BYTES).min(value.len());
        while end < value.len() && !value.is_char_boundary(end) {
            end += 1;
        }
        let encoded = encoding.encode(&value[start..end]).map_err(|_| {
            invalid(format!(
                "{field} is not representable in the project code page"
            ))
        })?;
        if encoded.contains(&0) {
            return Err(invalid(format!("{field} encodes to a null byte")));
        }
        length = length
            .checked_add(encoded.len())
            .ok_or_else(|| invalid(format!("{field} encoded size overflow")))?;
        start = end;
    }
    check_limit("VBA string bytes", length, limits.max_string_bytes)?;
    Ok(length)
}

fn encoded_libid_reference_len(
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    field: &'static str,
) -> Result<usize, Error> {
    encoded_reference_value_len(value, encoding, limits, field, |bytes| {
        validate_libid_reference(bytes, field)
    })
}

fn encoded_project_reference_len(
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    field: &'static str,
) -> Result<usize, Error> {
    encoded_reference_value_len(value, encoding, limits, field, |bytes| {
        validate_project_reference(bytes, field)
    })
}

fn encode_libid_reference_bytes(
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    field: &'static str,
) -> Result<Vec<u8>, Error> {
    let bytes = encode_reference_bytes(value, encoding, limits, field)?;
    validate_libid_reference(&bytes, field)?;
    Ok(bytes)
}

fn encode_project_reference_bytes(
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    field: &'static str,
) -> Result<Vec<u8>, Error> {
    let bytes = encode_reference_bytes(value, encoding, limits, field)?;
    validate_project_reference(&bytes, field)?;
    Ok(bytes)
}

fn validate_libid_reference(bytes: &[u8], field: &'static str) -> Result<(), Error> {
    let mut fields = bytes.splitn(5, |byte| *byte == b'#');
    let Some(first) = fields.next() else {
        return Err(invalid(format!("{field} is not a LibidReference")));
    };
    if first.len() < 4
        || first.get(..2) != Some(b"*\\".as_slice())
        || !matches!(first.get(2), Some(b'G' | b'H'))
        || !valid_guid(first.get(3..).unwrap_or_default())
    {
        return Err(invalid(format!("{field} is not a LibidReference")));
    }
    let Some(version) = fields.next() else {
        return Err(invalid(format!("{field} is not a LibidReference")));
    };
    let mut version_fields = version.split(|byte| *byte == b'.');
    let Some(major) = version_fields.next() else {
        return Err(invalid(format!("{field} is not a LibidReference")));
    };
    let Some(minor) = version_fields.next() else {
        return Err(invalid(format!("{field} is not a LibidReference")));
    };
    if version_fields.next().is_some() {
        return Err(invalid(format!("{field} is not a LibidReference")));
    }
    let Some(lcid) = fields.next() else {
        return Err(invalid(format!("{field} is not a LibidReference")));
    };
    let Some(path) = fields.next() else {
        return Err(invalid(format!("{field} is not a LibidReference")));
    };
    let Some(reg_name) = fields.next() else {
        return Err(invalid(format!("{field} is not a LibidReference")));
    };
    if !valid_hex(major, 1, 4)
        || !valid_hex(minor, 1, 4)
        || !valid_hex(lcid, 1, 8)
        || path.iter().any(|byte| *byte == 0 || *byte == 0x23)
        || reg_name.len() > 255
        || reg_name.contains(&0)
    {
        return Err(invalid(format!("{field} is not a LibidReference")));
    }
    Ok(())
}

fn validate_project_reference(bytes: &[u8], field: &'static str) -> Result<(), Error> {
    if bytes.len() < 3
        || bytes.get(..2) != Some(b"*\\".as_slice())
        || !matches!(bytes.get(2), Some(b'A' | b'B' | b'C' | b'D'))
        || bytes[3..].contains(&0)
    {
        return Err(invalid(format!("{field} is not a ProjectReference")));
    }
    Ok(())
}

fn valid_hex(bytes: &[u8], minimum: usize, maximum: usize) -> bool {
    (minimum..=maximum).contains(&bytes.len()) && bytes.iter().all(u8::is_ascii_hexdigit)
}

fn valid_guid(bytes: &[u8]) -> bool {
    bytes.len() == 38
        && bytes[0] == b'{'
        && bytes[9] == b'-'
        && bytes[14] == b'-'
        && bytes[19] == b'-'
        && bytes[24] == b'-'
        && bytes[37] == b'}'
        && bytes[1..9].iter().all(u8::is_ascii_hexdigit)
        && bytes[10..14].iter().all(u8::is_ascii_hexdigit)
        && bytes[15..19].iter().all(u8::is_ascii_hexdigit)
        && bytes[20..24].iter().all(u8::is_ascii_hexdigit)
        && bytes[25..37].iter().all(u8::is_ascii_hexdigit)
}

fn encode_control_reference(
    output: &mut Vec<u8>,
    control: &ControlReference,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<(), Error> {
    let twiddled = encode_libid_reference_bytes(
        &control.libid_twiddled,
        encoding,
        limits,
        "REFERENCECONTROL LibidTwiddled",
    )?;
    append_bounded(output, &REFERENCE_CONTROL_ID.to_le_bytes(), limits)?;
    let twiddled_size = 4usize
        .checked_add(twiddled.len())
        .and_then(|value| value.checked_add(4))
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| invalid("REFERENCECONTROL twiddled size overflow"))?;
    let twiddled_size = u32::try_from(twiddled_size)
        .map_err(|_| invalid("REFERENCECONTROL twiddled size exceeds u32"))?;
    append_bounded(output, &twiddled_size.to_le_bytes(), limits)?;
    push_length_prefixed_bounded(output, &twiddled, limits)?;
    append_bounded(output, &0u32.to_le_bytes(), limits)?;
    append_bounded(output, &0u16.to_le_bytes(), limits)?;

    let extended = control
        .extended
        .as_ref()
        .ok_or_else(|| invalid("REFERENCECONTROL requires extended type-library metadata"))?;
    let name = extended.name.as_deref();
    let extended_libid = extended.libid.as_str();
    let guid = extended.original_type_library;
    let cookie = extended.cookie;
    if let Some(name) = name {
        validate_reference_identifier(name, ReferenceNameKind::Library, "REFERENCENAME")?;
        push_string_pair(
            output,
            REFERENCE_NAME_ID,
            REFERENCE_NAME_RESERVED,
            name,
            encoding,
            limits,
            limits.max_string_bytes,
            "REFERENCECONTROL extended name",
        )?;
    }
    append_bounded(output, &REFERENCE_CONTROL_RESERVED.to_le_bytes(), limits)?;
    let extended_bytes = encode_libid_reference_bytes(
        extended_libid,
        encoding,
        limits,
        "REFERENCECONTROL LibidExtended",
    )?;
    let extended_size = 4usize
        .checked_add(extended_bytes.len())
        .and_then(|value| value.checked_add(4))
        .and_then(|value| value.checked_add(2))
        .and_then(|value| value.checked_add(16))
        .and_then(|value| value.checked_add(4))
        .ok_or_else(|| invalid("REFERENCECONTROL extended size overflow"))?;
    let extended_size = u32::try_from(extended_size)
        .map_err(|_| invalid("REFERENCECONTROL extended size exceeds u32"))?;
    append_bounded(output, &extended_size.to_le_bytes(), limits)?;
    push_length_prefixed_bounded(output, &extended_bytes, limits)?;
    append_bounded(output, &0u32.to_le_bytes(), limits)?;
    append_bounded(output, &0u16.to_le_bytes(), limits)?;
    append_bounded(output, &guid, limits)?;
    append_bounded(output, &cookie.to_le_bytes(), limits)?;
    Ok(())
}

fn push_record(output: &mut Vec<u8>, id: u16, value: &[u8]) -> Result<(), Error> {
    let Ok(length) = u32::try_from(value.len()) else {
        return Err(invalid("dir-stream record length exceeds u32"));
    };
    output.extend_from_slice(&id.to_le_bytes());
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(value);
    Ok(())
}

fn push_record_bounded(
    output: &mut Vec<u8>,
    id: u16,
    value: &[u8],
    limits: &Limits,
) -> Result<(), Error> {
    let added = value
        .len()
        .checked_add(6)
        .ok_or_else(|| invalid("dir-stream record size overflow"))?;
    check_append_limit(output, added, limits)?;
    push_record(output, id, value)
}

fn append_bounded(output: &mut Vec<u8>, value: &[u8], limits: &Limits) -> Result<(), Error> {
    check_append_limit(output, value.len(), limits)?;
    output.extend_from_slice(value);
    Ok(())
}

fn check_append_limit(output: &[u8], added: usize, limits: &Limits) -> Result<(), Error> {
    let new_len = output
        .len()
        .checked_add(added)
        .ok_or_else(|| invalid("dir-stream size overflow"))?;
    check_limit(
        "decompressed VBA stream bytes",
        new_len,
        limits.max_decompressed_stream_bytes,
    )
}

fn push_string_pair(
    output: &mut Vec<u8>,
    id: u16,
    reserved: u16,
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    protocol_maximum: usize,
    field: &'static str,
) -> Result<(), Error> {
    check_limit(
        "VBA input string bytes",
        value.len(),
        limits.max_string_bytes.saturating_mul(4),
    )?;
    let mbcs_len = encoded_mbcs_len(value, encoding, limits, field)?;
    check_protocol_length(field, mbcs_len, protocol_maximum)?;
    let first_record = mbcs_len
        .checked_add(6)
        .ok_or_else(|| invalid(format!("{field} MBCS record size overflow")))?;
    // Preserve the writer's record-order error reporting while ensuring the
    // full MBCS value is never materialized before this aggregate check.
    check_append_limit(output, first_record, limits)?;
    let unicode_len = value
        .encode_utf16()
        .count()
        .checked_mul(2)
        .ok_or_else(|| invalid(format!("{field} UTF-16 length overflow")))?;
    check_limit("VBA string bytes", unicode_len, limits.max_string_bytes)?;
    check_protocol_length(field, unicode_len, protocol_maximum.saturating_mul(2))?;
    let total = first_record
        .checked_add(2)
        .and_then(|length| length.checked_add(4))
        .and_then(|length| length.checked_add(unicode_len))
        .ok_or_else(|| invalid(format!("{field} encoded size overflow")))?;
    check_append_limit(output, total, limits)?;
    let mbcs = encode_mbcs(value, encoding, field)?;
    push_record_bounded(output, id, &mbcs, limits)?;
    append_bounded(output, &reserved.to_le_bytes(), limits)?;
    let unicode = encode_utf16(value, field)?;
    push_length_prefixed_bounded(output, &unicode, limits)
}

fn push_mbcs_pair(
    output: &mut Vec<u8>,
    id: u16,
    reserved: u16,
    value: &str,
    encoding: Mbcs,
    limits: &Limits,
    protocol_maximum: usize,
    field: &'static str,
) -> Result<(), Error> {
    check_limit(
        "VBA input string bytes",
        value.len(),
        limits.max_string_bytes.saturating_mul(4),
    )?;
    let mbcs_len = encoded_mbcs_len(value, encoding, limits, field)?;
    check_protocol_length(field, mbcs_len, protocol_maximum)?;
    let first_record = mbcs_len
        .checked_add(6)
        .ok_or_else(|| invalid(format!("{field} MBCS record size overflow")))?;
    check_append_limit(output, first_record, limits)?;
    let second_record = mbcs_len
        .checked_add(4)
        .ok_or_else(|| invalid(format!("{field} MBCS string size overflow")))?;
    let total = first_record
        .checked_add(2)
        .and_then(|length| length.checked_add(second_record))
        .ok_or_else(|| invalid(format!("{field} MBCS pair size overflow")))?;
    check_append_limit(output, total, limits)?;
    let mbcs = encode_mbcs(value, encoding, field)?;
    push_record_bounded(output, id, &mbcs, limits)?;
    append_bounded(output, &reserved.to_le_bytes(), limits)?;
    push_length_prefixed_bounded(output, &mbcs, limits)
}

fn push_length_prefixed(output: &mut Vec<u8>, value: &[u8]) -> Result<(), Error> {
    let Ok(length) = u32::try_from(value.len()) else {
        return Err(invalid("dir-stream string length exceeds u32"));
    };
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(value);
    Ok(())
}

fn push_length_prefixed_bounded(
    output: &mut Vec<u8>,
    value: &[u8],
    limits: &Limits,
) -> Result<(), Error> {
    let added = value
        .len()
        .checked_add(4)
        .ok_or_else(|| invalid("dir-stream string size overflow"))?;
    check_append_limit(output, added, limits)?;
    push_length_prefixed(output, value)
}

fn check_protocol_length(field: &'static str, actual: usize, maximum: usize) -> Result<(), Error> {
    if actual > maximum {
        return Err(invalid(format!(
            "{field} length {actual} exceeds {maximum}"
        )));
    }
    Ok(())
}

pub(crate) fn encode_mbcs(
    value: &str,
    encoding: Mbcs,
    field: &'static str,
) -> Result<Vec<u8>, Error> {
    if value.contains('\0') {
        return Err(invalid(format!("{field} contains a null character")));
    }
    let Ok(encoded) = encoding.encode(value) else {
        return Err(invalid(format!(
            "{field} is not representable in the project code page"
        )));
    };
    if encoded.contains(&0) {
        return Err(invalid(format!("{field} encodes to a null byte")));
    }
    Ok(encoded.into_owned())
}

fn encode_utf16(value: &str, field: &'static str) -> Result<Vec<u8>, Error> {
    if value.contains('\0') {
        return Err(invalid(format!("{field} contains a null character")));
    }
    Ok(value.encode_utf16().flat_map(u16::to_le_bytes).collect())
}

fn parse_project_information(
    data: &[u8],
    limits: &Limits,
) -> Result<(usize, Mbcs, Vec<u8>), Error> {
    let mut reader = Reader::new(data, 0);
    reader.expect_sized_u32(PROJECT_SYS_KIND_ID, FIXED_U32_SIZE)?;
    if reader.read_u32()? > 3 {
        return Err(invalid("PROJECTSYSKIND contains an unknown platform"));
    }
    if reader.peek_u16() == Some(PROJECT_COMPAT_VERSION_ID) {
        reader.expect_sized_u32(PROJECT_COMPAT_VERSION_ID, FIXED_U32_SIZE)?;
        reader.read_u32()?;
    }
    reader.expect_sized_u32(PROJECT_LCID_ID, FIXED_U32_SIZE)?;
    reader.expect_u32(DEFAULT_PROJECT_LCID, "PROJECTLCID value")?;
    reader.expect_sized_u32(PROJECT_LCID_INVOKE_ID, FIXED_U32_SIZE)?;
    reader.expect_u32(DEFAULT_PROJECT_LCID, "PROJECTLCIDINVOKE value")?;
    reader.expect_sized_u32(PROJECT_CODEPAGE_ID, FIXED_U16_SIZE)?;
    let code_page = reader.read_u16()?;

    reader.expect_id(PROJECT_NAME_ID)?;
    let project_name = reader
        .length_prefixed_bounded(limits, 128, "PROJECTNAME")?
        .to_vec();
    if project_name.is_empty() {
        return Err(invalid("PROJECTNAME must not be empty"));
    }
    let encoding = Mbcs::new(u32::from(code_page)).ok_or(Error::UnsupportedCodePage(code_page))?;
    decode_mbcs(&project_name, encoding, "PROJECTNAME")?;

    reader.expect_id(PROJECT_DOC_STRING_ID)?;
    reader.string_pair_bounded(
        encoding,
        PROJECT_DOC_STRING_RESERVED,
        "PROJECTDOCSTRING",
        limits,
        2_000,
    )?;
    reader.expect_id(PROJECT_HELP_FILE_PATH_ID)?;
    reader.mbcs_pair_bounded(
        encoding,
        PROJECT_HELP_FILE_PATH_RESERVED,
        "PROJECTHELPFILEPATH",
        limits,
        260,
    )?;
    reader.expect_sized_u32(PROJECT_HELP_CONTEXT_ID, FIXED_U32_SIZE)?;
    reader.read_u32()?;
    reader.expect_sized_u32(PROJECT_LIB_FLAGS_ID, FIXED_U32_SIZE)?;
    reader.expect_u32(0, "PROJECTLIBFLAGS value")?;
    reader.expect_id(PROJECT_VERSION_ID)?;
    let _project_version_reserved = reader.read_u32()?;
    reader.read_u32()?;
    reader.read_u16()?;
    if reader.peek_u16() == Some(PROJECT_CONSTANTS_ID) {
        reader.expect_id(PROJECT_CONSTANTS_ID)?;
        reader.string_pair_bounded(
            encoding,
            PROJECT_CONSTANTS_RESERVED,
            "PROJECTCONSTANTS",
            limits,
            1_015,
        )?;
    }
    Ok((reader.position, encoding, project_name))
}

fn find_references_and_modules(
    data: &[u8],
    encoding: Mbcs,
    limits: &Limits,
    require_end: bool,
) -> Result<(Vec<Reference>, Vec<Module>), Error> {
    let mut reader = Reader::new(data, 0);
    let mut references = Vec::new();
    loop {
        let id = reader
            .peek_u16()
            .ok_or_else(|| invalid("missing PROJECTMODULES record"))?;
        if id == PROJECT_MODULES_ID {
            break;
        }
        check_limit(
            "VBA reference count",
            references.len().saturating_add(1),
            limits.max_references,
        )?;
        references.push(parse_reference(&mut reader, encoding, limits)?);
    }

    reader.expect_sized_u32(PROJECT_MODULES_ID, FIXED_U16_SIZE)?;
    let count = usize::from(reader.read_u16()?);
    check_limit("VBA module count", count, limits.max_modules)?;
    reader.expect_sized_u32(PROJECT_COOKIE_ID, FIXED_U16_SIZE)?;
    reader.read_u16()?;
    let modules = parse_modules(&mut reader, count, encoding, limits)?;
    reader.expect_id(DIR_TERMINATOR_ID)?;
    let _dir_terminator_reserved = reader.read_u32()?;
    if require_end && reader.position != data.len() {
        return Err(invalid(
            "dir stream has trailing bytes after its terminator",
        ));
    }
    Ok((references, modules))
}

fn parse_reference(
    reader: &mut Reader<'_>,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<Reference, Error> {
    let start = reader.position;
    let name = if reader.peek_u16() == Some(REFERENCE_NAME_ID) {
        Some(parse_reference_name(reader, encoding, limits)?)
    } else {
        None
    };
    let id = reader
        .peek_u16()
        .ok_or_else(|| invalid("reference record is truncated"))?;
    let kind = match id {
        REFERENCE_REGISTERED_ID => ReferenceKind::Registered {
            libid: parse_registered_reference(reader, encoding, limits)?,
        },
        REFERENCE_PROJECT_ID => parse_project_reference(reader, encoding, limits)?,
        REFERENCE_CONTROL_ID => {
            let control = parse_control_reference(reader, encoding, limits)?;
            ReferenceKind::Control {
                libid_twiddled: control.libid_twiddled,
                extended: control.extended,
            }
        },
        REFERENCE_ORIGINAL_ID => parse_original_reference(reader, encoding, limits)?,
        other => {
            reader.expect_id(other)?;
            let payload = reader.length_prefixed(limits)?.to_vec();
            ReferenceKind::Opaque { id: other, payload }
        },
    };
    let wire = reader
        .data
        .get(start..reader.position)
        .ok_or_else(|| invalid("reference source span is out of bounds"))?
        .to_vec();
    Ok(Reference {
        name,
        kind,
        wire: Some(wire),
    })
}

fn parse_reference_name(
    reader: &mut Reader<'_>,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<String, Error> {
    reader.expect_id(REFERENCE_NAME_ID)?;
    let mbcs = decode_mbcs(reader.length_prefixed(limits)?, encoding, "REFERENCENAME")?;
    let _reserved = reader.read_u16()?;
    let unicode = decode_utf16(reader.length_prefixed(limits)?, "REFERENCENAMEUNICODE")?;
    if unicode != mbcs {
        return Err(invalid("REFERENCENAMEUNICODE does not match REFERENCENAME"));
    }
    Ok(unicode)
}

fn parse_registered_reference(
    reader: &mut Reader<'_>,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<String, Error> {
    reader.expect_id(REFERENCE_REGISTERED_ID)?;
    let _size = reader.read_u32()?;
    let libid = decode_mbcs(
        reader.length_prefixed(limits)?,
        encoding,
        "REFERENCEREGISTERED Libid",
    )?;
    validate_libid_reference(libid.as_bytes(), "REFERENCEREGISTERED Libid")?;
    let _reserved1 = reader.read_u32()?;
    let _reserved2 = reader.read_u16()?;
    Ok(libid)
}

fn parse_project_reference(
    reader: &mut Reader<'_>,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<ReferenceKind, Error> {
    reader.expect_id(REFERENCE_PROJECT_ID)?;
    let _size = reader.read_u32()?;
    let absolute = decode_mbcs(
        reader.length_prefixed(limits)?,
        encoding,
        "REFERENCEPROJECT LibidAbsolute",
    )?;
    let relative = decode_mbcs(
        reader.length_prefixed(limits)?,
        encoding,
        "REFERENCEPROJECT LibidRelative",
    )?;
    validate_project_reference(absolute.as_bytes(), "REFERENCEPROJECT LibidAbsolute")?;
    validate_project_reference(relative.as_bytes(), "REFERENCEPROJECT LibidRelative")?;
    let major_version = reader.read_u32()?;
    let minor_version = reader.read_u16()?;
    Ok(ReferenceKind::Project {
        libid_absolute: absolute,
        libid_relative: relative,
        major_version,
        minor_version,
    })
}

fn parse_control_reference(
    reader: &mut Reader<'_>,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<ControlReference, Error> {
    reader.expect_id(REFERENCE_CONTROL_ID)?;
    let _size_twiddled = reader.read_u32()?;
    let libid_twiddled = decode_mbcs(
        reader.length_prefixed(limits)?,
        encoding,
        "REFERENCECONTROL LibidTwiddled",
    )?;
    validate_libid_reference(libid_twiddled.as_bytes(), "REFERENCECONTROL LibidTwiddled")?;
    let _reserved1 = reader.read_u32()?;
    let _reserved2 = reader.read_u16()?;
    let name = if reader.peek_u16() == Some(REFERENCE_NAME_ID) {
        Some(parse_reference_name(reader, encoding, limits)?)
    } else {
        None
    };
    let _reserved3 = reader.read_u16()?;
    let _size_extended = reader.read_u32()?;
    let libid_extended = decode_mbcs(
        reader.length_prefixed(limits)?,
        encoding,
        "REFERENCECONTROL LibidExtended",
    )?;
    validate_libid_reference(libid_extended.as_bytes(), "REFERENCECONTROL LibidExtended")?;
    let _reserved4 = reader.read_u32()?;
    let _reserved5 = reader.read_u16()?;
    let guid_bytes = reader.read_bytes(16)?;
    let mut original_type_library = [0u8; 16];
    original_type_library.copy_from_slice(guid_bytes);
    let cookie = reader.read_u32()?;
    Ok(ControlReference::new(
        libid_twiddled,
        Some(ExtendedReference::new(
            name,
            libid_extended,
            original_type_library,
            cookie,
        )),
    ))
}

fn parse_original_reference(
    reader: &mut Reader<'_>,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<ReferenceKind, Error> {
    reader.expect_id(REFERENCE_ORIGINAL_ID)?;
    let libid_original = decode_mbcs(
        reader.length_prefixed(limits)?,
        encoding,
        "REFERENCEORIGINAL LibidOriginal",
    )?;
    validate_libid_reference(libid_original.as_bytes(), "REFERENCEORIGINAL LibidOriginal")?;
    let control = parse_control_reference(reader, encoding, limits)?;
    Ok(ReferenceKind::Original {
        libid_original,
        control,
    })
}

fn parse_modules(
    reader: &mut Reader<'_>,
    count: usize,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<Vec<Module>, Error> {
    const MODULE_MINIMUM_BYTES: usize = 6 + 12 + 12 + 10 + 10 + 8 + 6 + 6;
    let minimum_bytes = count
        .checked_mul(MODULE_MINIMUM_BYTES)
        .and_then(|value| value.checked_add(6))
        .ok_or_else(|| invalid("dir module minimum size overflows usize"))?;
    let remaining = reader.data.len().saturating_sub(reader.position);
    if remaining < minimum_bytes {
        return Err(invalid(format!(
            "dir stream is shorter than its {count} module records"
        )));
    }
    let mut modules = Vec::new();
    modules
        .try_reserve(count)
        .map_err(|_| invalid("dir module allocation failed"))?;
    for _ in 0..count {
        reader.expect_id(MODULE_NAME_ID)?;
        let name_bytes = reader.length_prefixed(limits)?;
        let mbcs_name = decode_mbcs(name_bytes, encoding, "MODULENAME")?;
        validate_module_identifier(&mbcs_name, "MODULENAME")?;

        let name = if reader.peek_u16() == Some(MODULE_NAME_UNICODE_ID) {
            reader.expect_id(MODULE_NAME_UNICODE_ID)?;
            let unicode = decode_utf16(reader.length_prefixed(limits)?, "MODULENAMEUNICODE")?;
            if unicode != mbcs_name {
                return Err(invalid(
                    "MODULENAMEUNICODE does not match the MBCS module name",
                ));
            }
            unicode
        } else {
            mbcs_name
        };
        validate_module_identifier(&name, "MODULENAME")?;

        reader.expect_id(MODULE_STREAM_NAME_ID)?;
        let stream_name =
            reader.string_pair(encoding, STREAM_NAME_RESERVED, "MODULESTREAMNAME", limits)?;
        validate_stream_name(&stream_name)?;

        reader.expect_id(MODULE_DOC_STRING_ID)?;
        let _description = reader.string_pair(
            encoding,
            MODULE_DOC_STRING_RESERVED,
            "MODULEDOCSTRING",
            limits,
        )?;

        reader.expect_sized_u32(MODULE_OFFSET_ID, FIXED_U32_SIZE)?;
        let text_offset = reader.read_u32()?;
        reader.expect_sized_u32(MODULE_HELP_CONTEXT_ID, FIXED_U32_SIZE)?;
        let _help_context = reader.read_u32()?;
        reader.expect_sized_u32(MODULE_COOKIE_ID, FIXED_U16_SIZE)?;
        let _cookie = reader.read_u16()?;

        let kind = match reader.read_u16()? {
            MODULE_PROCEDURAL_ID => Kind::Procedural,
            MODULE_OTHER_ID => Kind::DocumentClassOrDesigner,
            id => return Err(invalid(format!("unexpected MODULETYPE id {id:#06x}"))),
        };
        let _module_type_reserved = reader.read_u32()?;

        let mut read_only = false;
        let mut private = false;
        if reader.peek_u16() == Some(MODULE_READ_ONLY_ID) {
            reader.read_u16()?;
            let _read_only_reserved = reader.read_u32()?;
            read_only = true;
        }
        if reader.peek_u16() == Some(MODULE_PRIVATE_ID) {
            reader.read_u16()?;
            let _private_reserved = reader.read_u32()?;
            private = true;
        }
        reader.expect_id(MODULE_TERMINATOR_ID)?;
        let _module_terminator_reserved = reader.read_u32()?;

        modules.push(Module {
            name,
            stream_name,
            text_offset,
            kind,
            read_only,
            private,
        });
    }
    Ok(modules)
}

fn validate_module_identifier(value: &str, field: &'static str) -> Result<(), Error> {
    if value.is_empty() {
        return Err(invalid(format!("{field} must not be empty")));
    }
    if value.chars().count() > MAX_MODULE_IDENTIFIER_CHARACTERS {
        return Err(invalid(format!(
            "{field} exceeds {MAX_MODULE_IDENTIFIER_CHARACTERS} characters"
        )));
    }
    if value.contains('\0') {
        return Err(invalid(format!("{field} contains a null character")));
    }
    Ok(())
}

fn validate_stream_name(value: &str) -> Result<(), Error> {
    let code_units = value.encode_utf16().count();
    if code_units == 0 || code_units > MAX_CFB_NAME_CODE_UNITS {
        return Err(invalid(format!(
            "MODULESTREAMNAME must contain 1 to {MAX_CFB_NAME_CODE_UNITS} UTF-16 code units"
        )));
    }
    if value
        .chars()
        .any(|character| character.is_control() || matches!(character, '/' | '\\' | ':' | '!'))
    {
        return Err(invalid(
            "MODULESTREAMNAME contains a forbidden CFB name character",
        ));
    }
    Ok(())
}

fn decode_mbcs(bytes: &[u8], encoding: Mbcs, field: &'static str) -> Result<String, Error> {
    if bytes.contains(&0) {
        return Err(invalid(format!("{field} contains a null byte")));
    }
    let Ok(decoded) = encoding.decode(bytes) else {
        return Err(invalid(format!(
            "{field} is invalid for the project code page"
        )));
    };
    Ok(decoded.into_owned())
}

pub(crate) fn decode_utf16(bytes: &[u8], field: &'static str) -> Result<String, Error> {
    if !bytes.len().is_multiple_of(2) {
        return Err(invalid(format!("{field} byte length is not even")));
    }
    let code_units = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
    let mut value = String::with_capacity(bytes.len() / 2);
    for decoded_unit in char::decode_utf16(code_units) {
        let Ok(character) = decoded_unit else {
            return Err(invalid(format!("{field} contains invalid UTF-16")));
        };
        if character == '\0' {
            return Err(invalid(format!("{field} contains a null character")));
        }
        value.push(character);
    }
    Ok(value)
}

fn read_u16_at(data: &[u8], position: usize) -> Option<u16> {
    let bytes = data.get(position..position.checked_add(2)?)?;
    Some(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32_at(data: &[u8], position: usize) -> Option<u32> {
    let bytes = data.get(position..position.checked_add(4)?)?;
    Some(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "test fixtures and assertions panic intentionally on failure"
    )]

    use super::*;

    fn push_record(bytes: &mut Vec<u8>, id: u16, value: &[u8]) {
        bytes.extend_from_slice(&id.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(value.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(value);
    }

    fn push_string_pair(bytes: &mut Vec<u8>, id: u16, value: &str, reserved: u16) {
        push_record(bytes, id, value.as_bytes());
        bytes.extend_from_slice(&reserved.to_le_bytes());
        let utf16: Vec<u8> = value.encode_utf16().flat_map(u16::to_le_bytes).collect();
        bytes.extend_from_slice(&u32::try_from(utf16.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(&utf16);
    }

    fn push_project_information(bytes: &mut Vec<u8>, name: &str) {
        push_record(bytes, PROJECT_SYS_KIND_ID, &1u32.to_le_bytes());
        push_record(bytes, PROJECT_LCID_ID, &DEFAULT_PROJECT_LCID.to_le_bytes());
        push_record(
            bytes,
            PROJECT_LCID_INVOKE_ID,
            &DEFAULT_PROJECT_LCID.to_le_bytes(),
        );
        push_record(bytes, PROJECT_CODEPAGE_ID, &1252u16.to_le_bytes());
        push_record(bytes, PROJECT_NAME_ID, name.as_bytes());
        push_string_pair(
            bytes,
            PROJECT_DOC_STRING_ID,
            "",
            PROJECT_DOC_STRING_RESERVED,
        );
        push_record(bytes, PROJECT_HELP_FILE_PATH_ID, &[]);
        bytes.extend_from_slice(&PROJECT_HELP_FILE_PATH_RESERVED.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        push_record(bytes, PROJECT_HELP_CONTEXT_ID, &0u32.to_le_bytes());
        push_record(bytes, PROJECT_LIB_FLAGS_ID, &0u32.to_le_bytes());
        bytes.extend_from_slice(&PROJECT_VERSION_ID.to_le_bytes());
        bytes.extend_from_slice(&FIXED_U32_SIZE.to_le_bytes());
        bytes.extend_from_slice(&1u32.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
    }

    fn literal_container(data: &[u8]) -> Vec<u8> {
        let mut encoded = vec![0x01];
        let mut chunk = Vec::new();
        for literals in data.chunks(8) {
            chunk.push(0);
            chunk.extend_from_slice(literals);
        }
        let header = 0xb000 | u16::try_from(chunk.len() - 1).unwrap();
        encoded.extend_from_slice(&header.to_le_bytes());
        encoded.extend_from_slice(&chunk);
        encoded
    }

    fn sample_dir() -> Vec<u8> {
        let mut bytes = Vec::new();
        push_project_information(&mut bytes, "Sample");
        push_record(&mut bytes, PROJECT_MODULES_ID, &1u16.to_le_bytes());
        push_record(&mut bytes, PROJECT_COOKIE_ID, &0xffffu16.to_le_bytes());
        push_record(&mut bytes, MODULE_NAME_ID, b"Module1");
        push_string_pair(
            &mut bytes,
            MODULE_STREAM_NAME_ID,
            "Module1",
            STREAM_NAME_RESERVED,
        );
        push_string_pair(
            &mut bytes,
            MODULE_DOC_STRING_ID,
            "",
            MODULE_DOC_STRING_RESERVED,
        );
        push_record(&mut bytes, MODULE_OFFSET_ID, &12u32.to_le_bytes());
        push_record(&mut bytes, MODULE_HELP_CONTEXT_ID, &0u32.to_le_bytes());
        push_record(&mut bytes, MODULE_COOKIE_ID, &0xffffu16.to_le_bytes());
        bytes.extend_from_slice(&MODULE_PROCEDURAL_ID.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&MODULE_READ_ONLY_ID.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&MODULE_PRIVATE_ID.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&MODULE_TERMINATOR_ID.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&DIR_TERMINATOR_ID.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        literal_container(&bytes)
    }

    #[test]
    fn parses_typed_module_directory() {
        let directory = Dir::parse(&sample_dir(), &Limits::default()).unwrap();
        assert_eq!(directory.page(), Mbcs::WINDOWS_1252);
        assert_eq!(directory.project_name(), "Sample");
        assert_eq!(directory.modules().len(), 1);
        let module = &directory.modules()[0];
        assert_eq!(module.name(), "Module1");
        assert_eq!(module.stream_name(), "Module1");
        assert_eq!(module.text_offset(), 12);
        assert_eq!(module.kind(), Kind::Procedural);
        assert!(module.is_read_only());
        assert!(module.is_private());
    }

    #[test]
    fn rejects_module_count_over_limit() {
        let limits = Limits {
            max_modules: 0,
            ..Limits::default()
        };
        assert!(matches!(
            Dir::parse(&sample_dir(), &limits),
            Err(Error::LimitExceeded { .. })
        ));
    }

    #[test]
    fn rejects_missing_top_level_directory_terminator() {
        let limits = Limits::default();
        let mut decompressed = codec::decode(&sample_dir(), &limits).unwrap();
        decompressed.truncate(decompressed.len() - 6);
        let malformed = codec::encode(&decompressed, &limits).unwrap();
        assert!(Dir::parse(&malformed, &limits).is_err());
    }

    #[test]
    fn source_bound_reference_name_edits_preserve_ignored_wire_fields() {
        let libid = b"*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##";
        let mut bytes = Vec::new();
        push_string_pair(&mut bytes, REFERENCE_NAME_ID, "stdole", 0x1234);
        bytes.extend_from_slice(&REFERENCE_REGISTERED_ID.to_le_bytes());
        bytes.extend_from_slice(&u32::MAX.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(libid.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(libid);
        bytes.extend_from_slice(&0xfeed_beefu32.to_le_bytes());
        bytes.extend_from_slice(&0xbeefu16.to_le_bytes());

        let mut reader = Reader::new(&bytes, 0);
        let source = parse_reference(&mut reader, Mbcs::WINDOWS_1252, &Limits::default()).unwrap();
        assert_eq!(source.raw(), Some(bytes.as_slice()));
        let unchanged = source
            .clone()
            .edit_name_source_bound(
                Some("stdole".to_owned()),
                Mbcs::WINDOWS_1252,
                &Limits::default(),
            )
            .unwrap();
        assert_eq!(unchanged.raw(), source.raw());
        assert_eq!(
            source.clone().with_name(Some("stdole".to_owned())).raw(),
            source.raw()
        );
        let edited = source
            .edit_name_source_bound(
                Some("stdole-edited".to_owned()),
                Mbcs::WINDOWS_1252,
                &Limits::default(),
            )
            .unwrap();
        let raw = edited.raw().unwrap();
        let reserved_offset = 2 + 4 + "stdole-edited".len();
        assert_eq!(
            &raw[reserved_offset..reserved_offset + 2],
            &0x1234u16.to_le_bytes()
        );
        assert!(
            raw.windows(4)
                .any(|field| field == 0xfeed_beefu32.to_le_bytes())
        );

        let mut encoded = Vec::new();
        encode_reference(
            &mut encoded,
            &edited,
            Mbcs::WINDOWS_1252,
            &Limits::default(),
        )
        .unwrap();
        assert_eq!(encoded, raw);
    }

    #[test]
    fn non_ascii_reference_preflight_preserves_encoding_and_limits() {
        let path = "é".repeat(600);
        let libid =
            format!("*\\G{{00000000-0000-0000-0000-000000000000}}#2.0#0#C:\\{path}#OLE Automation");
        let reference = Reference::registered("stdole", libid);
        let mut encoded = Vec::new();
        encode_reference(
            &mut encoded,
            &reference,
            Mbcs::WINDOWS_1252,
            &Limits::default(),
        )
        .unwrap();

        let mut reader = Reader::new(&encoded, 0);
        let parsed = parse_reference(&mut reader, Mbcs::WINDOWS_1252, &Limits::default()).unwrap();
        assert_eq!(parsed.name(), reference.name());
        assert_eq!(parsed.kind(), reference.kind());

        let limits = Limits {
            max_decompressed_stream_bytes: encoded.len() - 1,
            ..Limits::default()
        };
        let mut rejected = Vec::new();
        assert!(matches!(
            encode_reference(&mut rejected, &reference, Mbcs::WINDOWS_1252, &limits),
            Err(Error::LimitExceeded {
                limit: "decompressed VBA stream bytes",
                actual,
                maximum
            }) if actual == encoded.len() && maximum == encoded.len() - 1
        ));
        assert!(rejected.is_empty());
    }

    #[test]
    fn reader_ignores_reserved_words_and_accepts_hexdig_case_variants() {
        let limits = Limits::default();
        let mut decompressed = codec::decode(&sample_dir(), &limits).unwrap();

        let module_stream_id = MODULE_STREAM_NAME_ID.to_le_bytes();
        let stream_record = decompressed
            .windows(module_stream_id.len())
            .position(|window| window == module_stream_id)
            .unwrap();
        let stream_length_start = stream_record + 2;
        let stream_length = usize::try_from(u32::from_le_bytes(
            decompressed[stream_length_start..stream_length_start + 4]
                .try_into()
                .unwrap(),
        ))
        .unwrap();
        let stream_reserved = stream_record + 6 + stream_length;
        decompressed[stream_reserved..stream_reserved + 2]
            .copy_from_slice(&0xeeeeu16.to_le_bytes());

        let project_version_id = PROJECT_VERSION_ID.to_le_bytes();
        let version_record = decompressed
            .windows(project_version_id.len())
            .position(|window| window == project_version_id)
            .unwrap();
        decompressed[version_record + 2..version_record + 6]
            .copy_from_slice(&0xdead_beefu32.to_le_bytes());

        let terminator = decompressed.len() - 4;
        decompressed[terminator..].copy_from_slice(&0xfeed_beefu32.to_le_bytes());
        let compressed = codec::encode(&decompressed, &limits).unwrap();
        let directory = Dir::parse(&compressed, &limits).unwrap();
        assert_eq!(directory.modules()[0].stream_name(), "Module1");

        assert!(
            validate_libid_reference(
                b"*\\G{abcdefab-cdef-abcd-efab-cdefabcdefab}#a.b#c##",
                "test LibidReference"
            )
            .is_ok()
        );
    }

    #[test]
    fn unknown_reference_payloads_obey_string_limits() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0x7f00u16.to_le_bytes());
        bytes.extend_from_slice(&2u32.to_le_bytes());
        bytes.extend_from_slice(&[1, 2]);
        let limits = Limits {
            max_string_bytes: 1,
            ..Limits::default()
        };
        let mut reader = Reader::new(&bytes, 0);
        assert!(matches!(
            parse_reference(&mut reader, Mbcs::WINDOWS_1252, &limits),
            Err(Error::LimitExceeded {
                limit: "VBA string bytes",
                actual: 2,
                maximum: 1,
            })
        ));
    }
}
