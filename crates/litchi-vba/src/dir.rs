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

const STREAM_NAME_RESERVED: u16 = 0x0032;
const PROJECT_DOC_STRING_RESERVED: u16 = 0x0040;
const PROJECT_HELP_FILE_PATH_RESERVED: u16 = 0x003d;
const PROJECT_CONSTANTS_RESERVED: u16 = 0x003c;
const MODULE_DOC_STRING_RESERVED: u16 = 0x0048;
const FIXED_U32_SIZE: u32 = 4;
const FIXED_U16_SIZE: u32 = 2;
const DEFAULT_PROJECT_LCID: u32 = 0x0409;
const WRITE_COOKIE: u16 = 0xffff;

pub(crate) struct WriteProject<'a> {
    pub(crate) system_kind: u32,
    pub(crate) page: Mbcs,
    pub(crate) name: &'a str,
    pub(crate) description: &'a str,
    pub(crate) help_context: u32,
    pub(crate) version_major: u32,
    pub(crate) version_minor: u16,
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

pub(crate) fn encode_dir(project: &WriteProject<'_>, limits: &Limits) -> Result<Vec<u8>, Error> {
    check_limit(
        "VBA module count",
        project.modules.len(),
        limits.max_modules,
    )?;
    let module_count = u16::try_from(project.modules.len())
        .map_err(|_| invalid("VBA module count exceeds the dir-stream field"))?;
    let encoding = project.page;
    let mut output = Vec::new();

    push_record(
        &mut output,
        PROJECT_SYS_KIND_ID,
        &project.system_kind.to_le_bytes(),
    )?;
    push_record(
        &mut output,
        PROJECT_LCID_ID,
        &DEFAULT_PROJECT_LCID.to_le_bytes(),
    )?;
    push_record(
        &mut output,
        PROJECT_LCID_INVOKE_ID,
        &DEFAULT_PROJECT_LCID.to_le_bytes(),
    )?;
    push_record(
        &mut output,
        PROJECT_CODEPAGE_ID,
        &project.page.id16().to_le_bytes(),
    )?;

    let project_name = encode_mbcs(project.name, encoding, "PROJECTNAME")?;
    check_protocol_length("PROJECTNAME", project_name.len(), 128)?;
    push_record(&mut output, PROJECT_NAME_ID, &project_name)?;
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
    push_record(
        &mut output,
        PROJECT_HELP_CONTEXT_ID,
        &project.help_context.to_le_bytes(),
    )?;
    push_record(&mut output, PROJECT_LIB_FLAGS_ID, &0u32.to_le_bytes())?;
    output.extend_from_slice(&PROJECT_VERSION_ID.to_le_bytes());
    output.extend_from_slice(&FIXED_U32_SIZE.to_le_bytes());
    output.extend_from_slice(&project.version_major.to_le_bytes());
    output.extend_from_slice(&project.version_minor.to_le_bytes());
    check_limit(
        "decompressed VBA stream bytes",
        output.len(),
        limits.max_decompressed_stream_bytes,
    )?;

    push_record(&mut output, PROJECT_MODULES_ID, &module_count.to_le_bytes())?;
    push_record(&mut output, PROJECT_COOKIE_ID, &WRITE_COOKIE.to_le_bytes())?;
    for module in project.modules {
        encode_module(&mut output, module, encoding, limits)?;
        check_limit(
            "decompressed VBA stream bytes",
            output.len(),
            limits.max_decompressed_stream_bytes,
        )?;
    }
    output.extend_from_slice(&DIR_TERMINATOR_ID.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());

    check_limit(
        "decompressed VBA stream bytes",
        output.len(),
        limits.max_decompressed_stream_bytes,
    )?;
    codec::encode(&output, limits)
}

fn encode_module(
    output: &mut Vec<u8>,
    module: &WriteModule<'_>,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<(), Error> {
    let name = encode_mbcs(module.name, encoding, "MODULENAME")?;
    push_record(output, MODULE_NAME_ID, &name)?;
    let unicode_name = encode_utf16(module.name, "MODULENAMEUNICODE")?;
    push_record(output, MODULE_NAME_UNICODE_ID, &unicode_name)?;
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
    push_record(output, MODULE_OFFSET_ID, &0u32.to_le_bytes())?;
    push_record(
        output,
        MODULE_HELP_CONTEXT_ID,
        &module.help_context.to_le_bytes(),
    )?;
    push_record(output, MODULE_COOKIE_ID, &WRITE_COOKIE.to_le_bytes())?;
    let type_id = match module.kind {
        Kind::Procedural => MODULE_PROCEDURAL_ID,
        Kind::DocumentClassOrDesigner => MODULE_OTHER_ID,
    };
    output.extend_from_slice(&type_id.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    if module.read_only {
        output.extend_from_slice(&MODULE_READ_ONLY_ID.to_le_bytes());
        output.extend_from_slice(&0u32.to_le_bytes());
    }
    if module.private {
        output.extend_from_slice(&MODULE_PRIVATE_ID.to_le_bytes());
        output.extend_from_slice(&0u32.to_le_bytes());
    }
    output.extend_from_slice(&MODULE_TERMINATOR_ID.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    Ok(())
}

fn push_record(output: &mut Vec<u8>, id: u16, value: &[u8]) -> Result<(), Error> {
    let length =
        u32::try_from(value.len()).map_err(|_| invalid("dir-stream record length exceeds u32"))?;
    output.extend_from_slice(&id.to_le_bytes());
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(value);
    Ok(())
}

#[allow(clippy::too_many_arguments)]
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
    let mbcs = encode_mbcs(value, encoding, field)?;
    check_limit("VBA string bytes", mbcs.len(), limits.max_string_bytes)?;
    check_protocol_length(field, mbcs.len(), protocol_maximum)?;
    push_record(output, id, &mbcs)?;
    output.extend_from_slice(&reserved.to_le_bytes());
    let unicode = encode_utf16(value, field)?;
    check_limit("VBA string bytes", unicode.len(), limits.max_string_bytes)?;
    check_protocol_length(field, unicode.len(), protocol_maximum.saturating_mul(2))?;
    push_length_prefixed(output, &unicode)
}

#[allow(clippy::too_many_arguments)]
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
    let mbcs = encode_mbcs(value, encoding, field)?;
    check_limit("VBA string bytes", mbcs.len(), limits.max_string_bytes)?;
    check_protocol_length(field, mbcs.len(), protocol_maximum)?;
    push_record(output, id, &mbcs)?;
    output.extend_from_slice(&reserved.to_le_bytes());
    push_length_prefixed(output, &mbcs)
}

fn push_length_prefixed(output: &mut Vec<u8>, value: &[u8]) -> Result<(), Error> {
    let length =
        u32::try_from(value.len()).map_err(|_| invalid("dir-stream string length exceeds u32"))?;
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(value);
    Ok(())
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
    let encoded = encoding.encode(value).map_err(|_| {
        invalid(format!(
            "{field} is not representable in the project code page"
        ))
    })?;
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

/// Parsed metadata from an MS-OVBA `dir` stream.
#[derive(Debug, PartialEq, Eq)]
pub struct Dir {
    page: Mbcs,
    project_name: String,
    modules: Vec<Module>,
}

impl Dir {
    /// Parse a complete compressed `dir` stream.
    pub fn parse(compressed: &[u8], limits: &Limits) -> Result<Self, Error> {
        let decompressed = codec::decode(compressed, limits)?;
        Self::parse_decompressed(&decompressed, limits)
    }

    /// Checked page used by MBCS strings and module source.
    pub fn page(&self) -> Mbcs {
        self.page
    }

    /// Numeric page identifier stored in `PROJECTCODEPAGE`.
    pub fn page_id(&self) -> u16 {
        self.page.id16()
    }

    /// VBA project identifier from `PROJECTNAME`.
    pub fn project_name(&self) -> &str {
        &self.project_name
    }

    /// Module metadata in directory order.
    pub fn modules(&self) -> &[Module] {
        &self.modules
    }

    fn parse_decompressed(data: &[u8], limits: &Limits) -> Result<Self, Error> {
        let (information_end, page, project_name) = parse_project_information(data, limits)?;
        let project_name = decode_mbcs(&project_name, page, "PROJECTNAME")?;
        let modules = find_modules(&data[information_end..], page, limits)?;
        Ok(Self {
            page,
            project_name,
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
    pub fn name(&self) -> &str {
        &self.name
    }

    /// CFB stream name containing this module.
    pub fn stream_name(&self) -> &str {
        &self.stream_name
    }

    /// Byte offset at which compressed source begins in the module stream.
    pub fn text_offset(&self) -> u32 {
        self.text_offset
    }

    /// Broad module category from `MODULETYPE`.
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Whether `MODULEREADONLY` is present.
    pub fn is_read_only(&self) -> bool {
        self.read_only
    }

    /// Whether `MODULEPRIVATE` is present.
    pub fn is_private(&self) -> bool {
        self.private
    }
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
    reader.expect_u32(FIXED_U32_SIZE, "PROJECTVERSION reserved value")?;
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

impl<'a> Reader<'a> {
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
}

fn find_modules(data: &[u8], encoding: Mbcs, limits: &Limits) -> Result<Vec<Module>, Error> {
    let mut last_error = None;
    for position in 0..data.len().saturating_sub(15) {
        if read_u16_at(data, position) != Some(PROJECT_MODULES_ID)
            || read_u32_at(data, position + 2) != Some(FIXED_U16_SIZE)
            || read_u16_at(data, position + 8) != Some(PROJECT_COOKIE_ID)
            || read_u32_at(data, position + 10) != Some(FIXED_U16_SIZE)
        {
            continue;
        }
        let count = usize::from(read_u16_at(data, position + 6).unwrap_or(0));
        if let Err(error) = check_limit("VBA module count", count, limits.max_modules) {
            last_error = Some(error);
            continue;
        }
        let mut reader = Reader::new(data, position + 16);
        match parse_modules(&mut reader, count, encoding, limits) {
            Ok(modules) => {
                let terminator = reader
                    .expect_id(DIR_TERMINATOR_ID)
                    .and_then(|()| reader.expect_u32(0, "dir stream terminator reserved value"));
                match terminator {
                    Ok(()) => return Ok(modules),
                    Err(error) => last_error = Some(error),
                }
            },
            Err(error) => last_error = Some(error),
        }
    }
    Err(last_error.unwrap_or_else(|| invalid("missing PROJECTMODULES record")))
}

fn parse_modules(
    reader: &mut Reader<'_>,
    count: usize,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<Vec<Module>, Error> {
    let mut modules = Vec::with_capacity(count);
    for _ in 0..count {
        reader.expect_id(MODULE_NAME_ID)?;
        let name_bytes = reader.length_prefixed(limits)?;
        let mbcs_name = decode_mbcs(name_bytes, encoding, "MODULENAME")?;

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

        reader.expect_id(MODULE_STREAM_NAME_ID)?;
        let stream_name =
            reader.string_pair(encoding, STREAM_NAME_RESERVED, "MODULESTREAMNAME", limits)?;

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
        reader.expect_u32(0, "MODULETYPE reserved value")?;

        let mut read_only = false;
        let mut private = false;
        if reader.peek_u16() == Some(MODULE_READ_ONLY_ID) {
            reader.read_u16()?;
            reader.expect_u32(0, "MODULEREADONLY reserved value")?;
            read_only = true;
        }
        if reader.peek_u16() == Some(MODULE_PRIVATE_ID) {
            reader.read_u16()?;
            reader.expect_u32(0, "MODULEPRIVATE reserved value")?;
            private = true;
        }
        reader.expect_id(MODULE_TERMINATOR_ID)?;
        reader.expect_u32(0, "MODULE terminator reserved value")?;

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

fn decode_mbcs(bytes: &[u8], encoding: Mbcs, field: &'static str) -> Result<String, Error> {
    if bytes.contains(&0) {
        return Err(invalid(format!("{field} contains a null byte")));
    }
    let decoded = encoding
        .decode(bytes)
        .map_err(|_| invalid(format!("{field} is invalid for the project code page")))?;
    Ok(decoded.into_owned())
}

fn decode_utf16(bytes: &[u8], field: &'static str) -> Result<String, Error> {
    if !bytes.len().is_multiple_of(2) {
        return Err(invalid(format!("{field} byte length is not even")));
    }
    let code_units = bytes
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
    let mut value = String::with_capacity(bytes.len() / 2);
    for character in char::decode_utf16(code_units) {
        let character =
            character.map_err(|_| invalid(format!("{field} contains invalid UTF-16")))?;
        if character == '\0' {
            return Err(invalid(format!("{field} contains a null character")));
        }
        value.push(character);
    }
    Ok(value)
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
        let length = usize::try_from(self.read_u32()?)
            .map_err(|_| invalid("dir-stream string length does not fit usize"))?;
        check_limit("VBA string bytes", length, limits.max_string_bytes)?;
        self.read_bytes(length)
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
        reserved: u16,
        field: &'static str,
        limits: &Limits,
        protocol_maximum: usize,
    ) -> Result<String, Error> {
        let mbcs = decode_mbcs(
            self.length_prefixed_bounded(limits, protocol_maximum, field)?,
            encoding,
            field,
        )?;
        let actual_reserved = self.read_u16()?;
        if actual_reserved != reserved {
            return Err(invalid(format!(
                "{field} reserved id must be {reserved:#06x}, found {actual_reserved:#06x}"
            )));
        }
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
        reserved: u16,
        field: &'static str,
        limits: &Limits,
        protocol_maximum: usize,
    ) -> Result<String, Error> {
        let first_bytes = self.length_prefixed_bounded(limits, protocol_maximum, field)?;
        let first = decode_mbcs(first_bytes, encoding, field)?;
        let actual_reserved = self.read_u16()?;
        if actual_reserved != reserved {
            return Err(invalid(format!(
                "{field} reserved id must be {reserved:#06x}, found {actual_reserved:#06x}"
            )));
        }
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
    use super::*;

    fn push_record(bytes: &mut Vec<u8>, id: u16, value: &[u8]) {
        bytes.extend_from_slice(&id.to_le_bytes());
        bytes.extend_from_slice(&(value.len() as u32).to_le_bytes());
        bytes.extend_from_slice(value);
    }

    fn push_string_pair(bytes: &mut Vec<u8>, id: u16, value: &str, reserved: u16) {
        push_record(bytes, id, value.as_bytes());
        bytes.extend_from_slice(&reserved.to_le_bytes());
        let utf16: Vec<u8> = value.encode_utf16().flat_map(u16::to_le_bytes).collect();
        bytes.extend_from_slice(&(utf16.len() as u32).to_le_bytes());
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
}
