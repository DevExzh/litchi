//! Typed, inert MS-OVBA project authoring.

use super::directory::{
    DirectoryWriteModule, DirectoryWriteProject, encode_directory, encode_mbcs,
};
use super::{VbaError, VbaLimits, VbaModuleKind, check_limit, compress_container, invalid};
use crate::OleWriter;
use litchi_core::encoding::codepage_to_encoding;
use std::collections::HashSet;
use std::io::Cursor;

const VBA_STORAGE_NAME: &str = "VBA";
const DIR_STREAM_NAME: &str = "dir";
const PROJECT_STREAM_NAME: &str = "PROJECT";
const PROJECT_WM_STREAM_NAME: &str = "PROJECTwm";
const VERSION_PROJECT_STREAM_NAME: &str = "_VBA_PROJECT";
const VERSION_PROJECT_RESERVED: u16 = 0x61cc;
const VERSION_PROJECT_WRITE_VERSION: u16 = 0xffff;
const DEFAULT_CODE_PAGE: u16 = 1252;
const DEFAULT_PROJECT_VERSION_MAJOR: u32 = 1;
const DEFAULT_PROJECT_VERSION_MINOR: u16 = 0;
const MAX_CFB_NAME_CODE_UNITS: usize = 31;
const MAX_VBA_IDENTIFIER_CHARACTERS: usize = 255;
const DETERMINISTIC_OBFUSCATION_SEED: u8 = 0;
const ENCRYPTION_VERSION: u8 = 2;
const PROJECT_VERSION_COMPATIBLE_32: &str = "393222000";

/// Target platform stored in `PROJECTSYSKIND`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VbaPlatform {
    /// 16-bit Windows.
    Windows16,
    /// 32-bit Windows.
    Windows32,
    /// Classic Macintosh.
    Macintosh,
    /// 64-bit Windows.
    Windows64,
}

impl VbaPlatform {
    const fn system_kind(self) -> u32 {
        match self {
            Self::Windows16 => 0,
            Self::Windows32 => 1,
            Self::Macintosh => 2,
            Self::Windows64 => 3,
        }
    }
}

/// Project-level module category represented in the `PROJECT` stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VbaProjectModuleKind {
    /// A standard procedural module (`Module=`).
    Standard,
    /// A class module (`Class=`).
    Class,
    /// A host document module (`Document=`).
    Document {
        /// Automation server version written after the module name.
        type_library_version: u32,
    },
}

impl VbaProjectModuleKind {
    const fn directory_kind(self) -> VbaModuleKind {
        match self {
            Self::Standard => VbaModuleKind::Procedural,
            Self::Class | Self::Document { .. } => VbaModuleKind::DocumentClassOrDesigner,
        }
    }
}

/// Canonical project Automation type-library identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct VbaProjectId([u8; 16]);

impl VbaProjectId {
    /// The all-zero project identifier.
    pub const NIL: Self = Self([0; 16]);

    /// Construct an identifier from bytes in canonical textual GUID order.
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    /// Return bytes in canonical textual GUID order.
    pub const fn as_bytes(self) -> [u8; 16] {
        self.0
    }

    fn braced_uppercase(self) -> String {
        let bytes = self.0;
        format!(
            "{{{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            bytes[0],
            bytes[1],
            bytes[2],
            bytes[3],
            bytes[4],
            bytes[5],
            bytes[6],
            bytes[7],
            bytes[8],
            bytes[9],
            bytes[10],
            bytes[11],
            bytes[12],
            bytes[13],
            bytes[14],
            bytes[15],
        )
    }
}

/// One module to be written into an inert VBA project.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaModuleBuilder {
    name: String,
    stream_name: String,
    kind: VbaProjectModuleKind,
    source_body: String,
    description: String,
    help_context: u32,
    read_only: bool,
    private: bool,
}

impl VbaModuleBuilder {
    /// Create a standard procedural module.
    pub fn standard(name: impl Into<String>, source_body: impl Into<String>) -> Self {
        Self::new(name, source_body, VbaProjectModuleKind::Standard)
    }

    /// Create a class module.
    pub fn class(name: impl Into<String>, source_body: impl Into<String>) -> Self {
        Self::new(name, source_body, VbaProjectModuleKind::Class)
    }

    /// Create a document module.
    pub fn document(
        name: impl Into<String>,
        type_library_version: u32,
        source_body: impl Into<String>,
    ) -> Self {
        Self::new(
            name,
            source_body,
            VbaProjectModuleKind::Document {
                type_library_version,
            },
        )
    }

    fn new(
        name: impl Into<String>,
        source_body: impl Into<String>,
        kind: VbaProjectModuleKind,
    ) -> Self {
        let name = name.into();
        Self {
            stream_name: name.clone(),
            name,
            kind,
            source_body: source_body.into(),
            description: String::new(),
            help_context: 0,
            read_only: false,
            private: false,
        }
    }

    /// Override the CFB stream name. By default it equals the module name.
    pub fn with_stream_name(mut self, stream_name: impl Into<String>) -> Self {
        self.stream_name = stream_name.into();
        self
    }

    /// Set the module description.
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = description.into();
        self
    }

    /// Set the module Help context identifier.
    pub fn with_help_context(mut self, help_context: u32) -> Self {
        self.help_context = help_context;
        self
    }

    /// Mark the module read-only.
    pub fn read_only(mut self, read_only: bool) -> Self {
        self.read_only = read_only;
        self
    }

    /// Mark the module private to its project.
    pub fn private(mut self, private: bool) -> Self {
        self.private = private;
        self
    }

    /// Module identifier.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Project-level module category.
    pub fn kind(&self) -> VbaProjectModuleKind {
        self.kind
    }
}

/// Builder for a cache-free, inert MS-OVBA project.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProjectBuilder {
    name: String,
    project_id: VbaProjectId,
    platform: VbaPlatform,
    code_page: u16,
    description: String,
    help_context: i32,
    version_major: u32,
    version_minor: u16,
    modules: Vec<VbaModuleBuilder>,
}

impl VbaProjectBuilder {
    /// Create a project with Windows-1252 text and 32-bit Windows metadata.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            project_id: VbaProjectId::NIL,
            platform: VbaPlatform::Windows32,
            code_page: DEFAULT_CODE_PAGE,
            description: String::new(),
            help_context: 0,
            version_major: DEFAULT_PROJECT_VERSION_MAJOR,
            version_minor: DEFAULT_PROJECT_VERSION_MINOR,
            modules: Vec::new(),
        }
    }

    /// Set the project Automation type-library identifier.
    pub fn with_project_id(mut self, project_id: VbaProjectId) -> Self {
        self.project_id = project_id;
        self
    }

    /// Set the target platform.
    pub fn with_platform(mut self, platform: VbaPlatform) -> Self {
        self.platform = platform;
        self
    }

    /// Set the MBCS code page used for project text and module source.
    pub fn with_code_page(mut self, code_page: u16) -> Self {
        self.code_page = code_page;
        self
    }

    /// Set the project description.
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = description.into();
        self
    }

    /// Set the project Help context identifier.
    pub fn with_help_context(mut self, help_context: i32) -> Self {
        self.help_context = help_context;
        self
    }

    /// Set the project version mirrored into the `dir` stream.
    pub fn with_version(mut self, major: u32, minor: u16) -> Self {
        self.version_major = major;
        self.version_minor = minor;
        self
    }

    /// Append a module in directory order.
    pub fn add_module(&mut self, module: VbaModuleBuilder) -> &mut Self {
        self.modules.push(module);
        self
    }

    /// Append a module and return the builder.
    pub fn with_module(mut self, module: VbaModuleBuilder) -> Self {
        self.modules.push(module);
        self
    }

    /// Serialize and validate the project without executing or compiling source.
    pub fn build(&self, limits: &VbaLimits) -> Result<VbaProjectBinary, VbaError> {
        check_limit("VBA module count", self.modules.len(), limits.max_modules)?;
        let encoding = codepage_to_encoding(u32::from(self.code_page))
            .ok_or(VbaError::UnsupportedCodePage(self.code_page))?;
        validate_project_name(&self.name)?;
        validate_quoted_text(&self.description, "VBA project description")?;
        let encoded_name = encode_mbcs(&self.name, encoding, "PROJECTNAME")?;
        if encoded_name.len() > 128 {
            return Err(invalid("PROJECTNAME exceeds 128 encoded bytes"));
        }
        validate_project_modules(&self.modules)?;

        let mut encoded_modules = Vec::with_capacity(self.modules.len());
        let mut total_source_bytes = 0usize;
        for module in &self.modules {
            let source = module_source(module);
            let source = encode_mbcs(&source, encoding, "module source")?;
            total_source_bytes = total_source_bytes
                .checked_add(source.len())
                .ok_or_else(|| invalid("aggregate VBA source size overflow"))?;
            check_limit(
                "aggregate VBA module source bytes",
                total_source_bytes,
                limits.max_total_source_bytes,
            )?;
            let compressed_source = compress_container(&source, limits)?;
            encoded_modules.push(EncodedModule {
                stream_name: module.stream_name.clone(),
                compressed_source,
            });
        }

        let directory_modules: Vec<_> = self
            .modules
            .iter()
            .map(|module| DirectoryWriteModule {
                name: &module.name,
                stream_name: &module.stream_name,
                description: &module.description,
                help_context: module.help_context,
                kind: module.kind.directory_kind(),
                read_only: module.read_only,
                private: module.private,
            })
            .collect();
        let directory = encode_directory(
            &DirectoryWriteProject {
                system_kind: self.platform.system_kind(),
                code_page: self.code_page,
                name: &self.name,
                description: &self.description,
                help_context: u32::from_le_bytes(self.help_context.to_le_bytes()),
                version_major: self.version_major,
                version_minor: self.version_minor,
                modules: &directory_modules,
            },
            limits,
        )?;
        let project_stream = encode_project_stream(self, encoding, limits)?;
        let project_wm_stream = encode_project_wm_stream(&self.modules, encoding, limits)?;

        Ok(VbaProjectBinary {
            version_project_stream: version_project_stream(),
            directory_stream: directory,
            project_stream,
            project_wm_stream,
            modules: encoded_modules,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct EncodedModule {
    stream_name: String,
    compressed_source: Vec<u8>,
}

/// Fully serialized streams for a cache-free, inert MS-OVBA project.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProjectBinary {
    version_project_stream: Vec<u8>,
    directory_stream: Vec<u8>,
    project_stream: Vec<u8>,
    project_wm_stream: Vec<u8>,
    modules: Vec<EncodedModule>,
}

impl VbaProjectBinary {
    /// Write all project storages and streams into an existing CFB writer.
    pub fn write_into(
        &self,
        writer: &mut OleWriter,
        project_root_path: &[&str],
    ) -> Result<(), VbaError> {
        let mut vba_storage: Vec<String> = project_root_path
            .iter()
            .map(|component| (*component).to_owned())
            .collect();
        vba_storage.push(VBA_STORAGE_NAME.to_owned());
        create_storage(writer, &vba_storage)?;

        let mut project_path: Vec<String> = project_root_path
            .iter()
            .map(|component| (*component).to_owned())
            .collect();
        project_path.push(PROJECT_STREAM_NAME.to_owned());
        create_stream(writer, &project_path, &self.project_stream)?;
        project_path.pop();
        project_path.push(PROJECT_WM_STREAM_NAME.to_owned());
        create_stream(writer, &project_path, &self.project_wm_stream)?;

        let mut vba_path = vba_storage;
        vba_path.push(VERSION_PROJECT_STREAM_NAME.to_owned());
        create_stream(writer, &vba_path, &self.version_project_stream)?;
        vba_path.pop();
        vba_path.push(DIR_STREAM_NAME.to_owned());
        create_stream(writer, &vba_path, &self.directory_stream)?;
        vba_path.pop();
        for module in &self.modules {
            vba_path.push(module.stream_name.clone());
            create_stream(writer, &vba_path, &module.compressed_source)?;
            vba_path.pop();
        }
        Ok(())
    }

    /// Build a standalone CFB payload suitable for an OOXML `vbaProject.bin`.
    pub fn to_cfb_bytes(&self) -> Result<Vec<u8>, VbaError> {
        let mut writer = OleWriter::new();
        self.write_into(&mut writer, &[])?;
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }

    /// Compressed `dir` stream bytes.
    pub fn directory_stream(&self) -> &[u8] {
        &self.directory_stream
    }

    /// Uncompressed `PROJECT` stream bytes.
    pub fn project_stream(&self) -> &[u8] {
        &self.project_stream
    }

    /// `PROJECTwm` module-name map bytes.
    pub fn project_wm_stream(&self) -> &[u8] {
        &self.project_wm_stream
    }
}

fn create_storage(writer: &mut OleWriter, path: &[String]) -> Result<(), VbaError> {
    let components: Vec<&str> = path.iter().map(String::as_str).collect();
    writer.create_storage(&components)?;
    Ok(())
}

fn create_stream(writer: &mut OleWriter, path: &[String], data: &[u8]) -> Result<(), VbaError> {
    let components: Vec<&str> = path.iter().map(String::as_str).collect();
    writer.create_stream(&components, data)?;
    Ok(())
}

fn validate_project_name(name: &str) -> Result<(), VbaError> {
    if name.is_empty() {
        return Err(invalid("VBA project name must not be empty"));
    }
    if name.chars().count() > 128 {
        return Err(invalid("VBA project name exceeds 128 characters"));
    }
    if name
        .chars()
        .any(|character| character == '"' || character.is_control())
    {
        return Err(invalid(
            "VBA project name contains a quoted-string delimiter or control character",
        ));
    }
    Ok(())
}

fn validate_project_modules(modules: &[VbaModuleBuilder]) -> Result<(), VbaError> {
    let mut names = HashSet::with_capacity(modules.len());
    let mut stream_names = HashSet::with_capacity(modules.len());
    for module in modules {
        validate_module_identifier(&module.name)?;
        validate_stream_name(&module.stream_name)?;
        if module.source_body.contains('\0') {
            return Err(invalid(format!(
                "VBA module {} source contains a null character",
                module.name
            )));
        }
        if !names.insert(module.name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate VBA module name {}",
                module.name
            )));
        }
        if !stream_names.insert(module.stream_name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate VBA module stream name {}",
                module.stream_name
            )));
        }
    }
    Ok(())
}

fn validate_module_identifier(name: &str) -> Result<(), VbaError> {
    if name.chars().count() > MAX_VBA_IDENTIFIER_CHARACTERS {
        return Err(invalid(format!(
            "VBA module name exceeds {MAX_VBA_IDENTIFIER_CHARACTERS} characters"
        )));
    }
    let mut characters = name.chars();
    let first = characters
        .next()
        .ok_or_else(|| invalid("VBA module name must not be empty"))?;
    if !first.is_alphabetic() {
        return Err(invalid(
            "VBA module name must begin with an alphabetic character",
        ));
    }
    if characters.any(|character| !(character.is_alphanumeric() || character == '_')) {
        return Err(invalid(
            "VBA module name contains a non-identifier character",
        ));
    }
    Ok(())
}

fn validate_stream_name(name: &str) -> Result<(), VbaError> {
    let code_units = name.encode_utf16().count();
    if code_units == 0 || code_units > MAX_CFB_NAME_CODE_UNITS {
        return Err(invalid(format!(
            "VBA module stream name must contain 1 to {MAX_CFB_NAME_CODE_UNITS} UTF-16 code units"
        )));
    }
    if name
        .chars()
        .any(|character| character.is_control() || matches!(character, '/' | '\\' | ':' | '!'))
    {
        return Err(invalid(
            "VBA module stream name contains a forbidden CFB name character",
        ));
    }
    if [DIR_STREAM_NAME, VERSION_PROJECT_STREAM_NAME]
        .iter()
        .any(|reserved| reserved.eq_ignore_ascii_case(name))
        || name
            .get(..6)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("__SRP_"))
    {
        return Err(invalid("VBA module stream name is reserved"));
    }
    Ok(())
}

fn validate_quoted_text(value: &str, field: &'static str) -> Result<(), VbaError> {
    if value
        .chars()
        .any(|character| character == '"' || character.is_control())
    {
        return Err(invalid(format!(
            "{field} contains a quoted-string delimiter or control character"
        )));
    }
    Ok(())
}

fn module_source(module: &VbaModuleBuilder) -> String {
    let mut source = String::with_capacity(module.name.len() + module.source_body.len() + 32);
    source.push_str("Attribute VB_Name = \"");
    source.push_str(&module.name);
    source.push_str("\"\r\n");
    source.push_str(&module.source_body);
    source
}

fn encode_project_stream(
    project: &VbaProjectBuilder,
    encoding: &'static encoding_rs::Encoding,
    limits: &VbaLimits,
) -> Result<Vec<u8>, VbaError> {
    let project_id = project.project_id.braced_uppercase();
    let mut text = String::new();
    text.push_str("ID=\"");
    text.push_str(&project_id);
    text.push_str("\"\r\n");
    for module in &project.modules {
        match module.kind {
            VbaProjectModuleKind::Standard => text.push_str("Module="),
            VbaProjectModuleKind::Class => text.push_str("Class="),
            VbaProjectModuleKind::Document {
                type_library_version,
            } => {
                text.push_str("Document=");
                text.push_str(&module.name);
                text.push_str("/&H");
                text.push_str(&format!("{type_library_version:08X}"));
                text.push_str("\r\n");
                continue;
            },
        }
        text.push_str(&module.name);
        text.push_str("\r\n");
    }
    text.push_str("Name=\"");
    text.push_str(&project.name);
    text.push_str("\"\r\n");
    text.push_str("HelpContextID=\"");
    text.push_str(&project.help_context.to_string());
    text.push_str("\"\r\n");
    if !project.description.is_empty() {
        text.push_str("Description=\"");
        text.push_str(&project.description);
        text.push_str("\"\r\n");
    }
    text.push_str("VersionCompatible32=\"");
    text.push_str(PROJECT_VERSION_COMPATIBLE_32);
    text.push_str("\"\r\n");

    let protection = encrypt_project_data(&[0; 4], &project_id);
    let password = encrypt_project_data(&[0], &project_id);
    let visibility = encrypt_project_data(&[0xff], &project_id);
    text.push_str("CMG=\"");
    push_upper_hex(&mut text, &protection);
    text.push_str("\"\r\nDPB=\"");
    push_upper_hex(&mut text, &password);
    text.push_str("\"\r\nGC=\"");
    push_upper_hex(&mut text, &visibility);
    text.push_str("\"\r\n\r\n[Host Extender Info]\r\n");

    let encoded = encode_mbcs(&text, encoding, "PROJECT stream")?;
    check_limit(
        "decompressed VBA stream bytes",
        encoded.len(),
        limits.max_decompressed_stream_bytes,
    )?;
    Ok(encoded)
}

fn encode_project_wm_stream(
    modules: &[VbaModuleBuilder],
    encoding: &'static encoding_rs::Encoding,
    limits: &VbaLimits,
) -> Result<Vec<u8>, VbaError> {
    let mut output = Vec::new();
    for module in modules {
        let encoded = encode_mbcs(&module.name, encoding, "PROJECTwm module name")?;
        check_limit("VBA string bytes", encoded.len(), limits.max_string_bytes)?;
        output.extend_from_slice(&encoded);
        output.push(0);
        let unicode: Vec<u16> = module.name.encode_utf16().collect();
        let unicode_bytes = unicode
            .len()
            .checked_mul(2)
            .ok_or_else(|| invalid("PROJECTwm Unicode name length overflow"))?;
        check_limit("VBA string bytes", unicode_bytes, limits.max_string_bytes)?;
        for code_unit in unicode {
            output.extend_from_slice(&code_unit.to_le_bytes());
        }
        output.extend_from_slice(&0u16.to_le_bytes());
        check_limit(
            "decompressed VBA stream bytes",
            output.len(),
            limits.max_decompressed_stream_bytes,
        )?;
    }
    output.extend_from_slice(&0u16.to_le_bytes());
    check_limit(
        "decompressed VBA stream bytes",
        output.len(),
        limits.max_decompressed_stream_bytes,
    )?;
    Ok(output)
}

fn version_project_stream() -> Vec<u8> {
    let mut output = Vec::with_capacity(7);
    output.extend_from_slice(&VERSION_PROJECT_RESERVED.to_le_bytes());
    output.extend_from_slice(&VERSION_PROJECT_WRITE_VERSION.to_le_bytes());
    output.push(0);
    output.extend_from_slice(&0u16.to_le_bytes());
    output
}

fn encrypt_project_data(data: &[u8], project_id: &str) -> Vec<u8> {
    let seed = DETERMINISTIC_OBFUSCATION_SEED;
    let project_key = project_id
        .bytes()
        .fold(0u8, |sum, byte| sum.wrapping_add(byte));
    let version_encrypted = seed ^ ENCRYPTION_VERSION;
    let project_key_encrypted = seed ^ project_key;
    let mut output = Vec::with_capacity(7 + data.len());
    output.extend([seed, version_encrypted, project_key_encrypted]);

    let mut unencrypted_byte_1 = project_key;
    let mut encrypted_byte_1 = project_key_encrypted;
    let mut encrypted_byte_2 = version_encrypted;
    for byte in u32::try_from(data.len())
        .expect("project encryption input is bounded")
        .to_le_bytes()
        .into_iter()
        .chain(data.iter().copied())
    {
        let encrypted = byte ^ encrypted_byte_2.wrapping_add(unencrypted_byte_1);
        output.push(encrypted);
        encrypted_byte_2 = encrypted_byte_1;
        encrypted_byte_1 = encrypted;
        unencrypted_byte_1 = byte;
    }
    output
}

fn push_upper_hex(output: &mut String, bytes: &[u8]) {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    output.reserve(bytes.len() * 2);
    for &byte in bytes {
        output.push(char::from(HEX[usize::from(byte >> 4)]));
        output.push(char::from(HEX[usize::from(byte & 0x0f)]));
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::OleFile;

    fn sample_builder() -> VbaProjectBuilder {
        let mut builder = VbaProjectBuilder::new("SampleProject")
            .with_project_id(VbaProjectId::from_bytes([
                0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc, 0xde, 0xf0, 0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc,
                0xde, 0xf0,
            ]))
            .with_description("Inert test project")
            .with_version(7, 3);
        builder
            .add_module(
                VbaModuleBuilder::standard("Module1", "Sub Main()\r\nEnd Sub\r\n")
                    .read_only(true)
                    .private(true),
            )
            .add_module(VbaModuleBuilder::class(
                "Class1",
                "Private value As Long\r\n",
            ))
            .add_module(VbaModuleBuilder::document(
                "ThisDocument",
                0x0001_0000,
                "Private Sub Document_Open()\r\nEnd Sub\r\n",
            ));
        builder
    }

    #[test]
    fn writes_complete_cache_free_project_and_reopens_every_module() {
        let limits = VbaLimits::default();
        let binary = sample_builder().build(&limits).unwrap();
        assert_eq!(
            binary.version_project_stream,
            [0xcc, 0x61, 0xff, 0xff, 0, 0, 0]
        );
        assert_eq!(
            binary,
            sample_builder().build(&limits).unwrap(),
            "authoring must be deterministic"
        );

        let bytes = binary.to_cfb_bytes().unwrap();
        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        assert_eq!(
            ole.open_stream(&["VBA", VERSION_PROJECT_STREAM_NAME])
                .unwrap(),
            [0xcc, 0x61, 0xff, 0xff, 0, 0, 0]
        );
        assert!(
            ole.list_directory_entries(&["VBA"])
                .unwrap()
                .iter()
                .all(|entry| !entry.name.starts_with("__SRP_"))
        );

        let project = super::super::VbaProject::open(&mut ole, &[], &limits).unwrap();
        assert_eq!(project.name(), "SampleProject");
        assert_eq!(project.modules().len(), 3);
        assert_eq!(project.modules()[0].name(), "Module1");
        assert!(project.modules()[0].is_read_only());
        assert!(project.modules()[0].is_private());
        assert_eq!(
            project.modules()[0].source().text(),
            "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nEnd Sub\r\n"
        );
        assert_eq!(project.modules()[1].name(), "Class1");
        assert_eq!(
            project.modules()[1].kind(),
            VbaModuleKind::DocumentClassOrDesigner
        );
        assert_eq!(project.modules()[2].name(), "ThisDocument");

        let properties = project.project_properties().text();
        assert!(properties.contains("Module=Module1\r\n"));
        assert!(properties.contains("Class=Class1\r\n"));
        assert!(properties.contains("Document=ThisDocument/&H00010000\r\n"));
        assert!(properties.contains("Description=\"Inert test project\"\r\n"));
        assert!(properties.contains("\r\n[Host Extender Info]\r\n"));
        assert_encrypted_property_shape(properties, "CMG", 22);
        assert_encrypted_property_shape(properties, "DPB", 16);
        assert_encrypted_property_shape(properties, "GC", 16);
    }

    #[test]
    fn writes_nested_legacy_project_and_non_ascii_codepage_text() {
        let mut builder = VbaProjectBuilder::new("日本語")
            .with_code_page(932)
            .with_platform(VbaPlatform::Windows64);
        builder.add_module(VbaModuleBuilder::standard(
            "標準",
            "Sub 挨拶()\r\nMsgBox \"こんにちは\"\r\nEnd Sub\r\n",
        ));
        let binary = builder.build(&VbaLimits::default()).unwrap();

        let mut writer = OleWriter::new();
        binary
            .write_into(&mut writer, &["_VBA_PROJECT_CUR"])
            .unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        bytes.set_position(0);
        let mut ole = OleFile::open(bytes).unwrap();
        let project =
            super::super::VbaProject::open(&mut ole, &["_VBA_PROJECT_CUR"], &VbaLimits::default())
                .unwrap();
        assert_eq!(project.name(), "日本語");
        assert_eq!(project.code_page(), 932);
        assert!(project.modules()[0].source().text().contains("こんにちは"));
        assert!(!project.modules()[0].source().had_decode_errors());
    }

    #[test]
    fn project_wm_contains_ordered_mbcs_and_unicode_name_pairs() {
        let limits = VbaLimits::default();
        let mut builder = VbaProjectBuilder::new("Map").with_code_page(932);
        builder
            .add_module(VbaModuleBuilder::standard("標準", ""))
            .add_module(VbaModuleBuilder::class("Class1", ""));
        let binary = builder.build(&limits).unwrap();
        let encoding = codepage_to_encoding(932).unwrap();
        let mut expected = Vec::new();
        for name in ["標準", "Class1"] {
            expected.extend_from_slice(&encode_mbcs(name, encoding, "test").unwrap());
            expected.push(0);
            expected.extend(name.encode_utf16().flat_map(u16::to_le_bytes));
            expected.extend_from_slice(&0u16.to_le_bytes());
        }
        expected.extend_from_slice(&0u16.to_le_bytes());
        assert_eq!(binary.project_wm_stream(), expected);
    }

    #[test]
    fn rejects_invalid_names_codepages_text_and_resource_limits() {
        let limits = VbaLimits::default();
        let mut duplicate = VbaProjectBuilder::new("Project");
        duplicate
            .add_module(VbaModuleBuilder::standard("Module1", ""))
            .add_module(VbaModuleBuilder::class("module1", ""));
        assert!(duplicate.build(&limits).is_err());

        let mut reserved = VbaProjectBuilder::new("Project");
        reserved.add_module(VbaModuleBuilder::standard("Module1", "").with_stream_name("dir"));
        assert!(reserved.build(&limits).is_err());
        let mut cache_name = VbaProjectBuilder::new("Project");
        cache_name
            .add_module(VbaModuleBuilder::standard("Module1", "").with_stream_name("__SRP_0"));
        assert!(cache_name.build(&limits).is_err());

        let mut invalid_identifier = VbaProjectBuilder::new("Project");
        invalid_identifier.add_module(VbaModuleBuilder::standard("1Module", ""));
        assert!(invalid_identifier.build(&limits).is_err());

        assert!(
            VbaProjectBuilder::new("Project")
                .with_description("invalid \"description")
                .build(&limits)
                .is_err()
        );
        let mut unrepresentable = VbaProjectBuilder::new("Project");
        unrepresentable.add_module(VbaModuleBuilder::standard("Module1", "MsgBox \"🙂\""));
        assert!(unrepresentable.build(&limits).is_err());
        assert!(matches!(
            VbaProjectBuilder::new("Project")
                .with_code_page(42)
                .build(&limits),
            Err(VbaError::UnsupportedCodePage(42))
        ));

        let no_modules = VbaLimits {
            max_modules: 0,
            ..limits
        };
        assert!(duplicate.build(&no_modules).is_err());
        let no_source = VbaLimits {
            max_total_source_bytes: 0,
            ..limits
        };
        let mut one_module = VbaProjectBuilder::new("Project");
        one_module.add_module(VbaModuleBuilder::standard("Module1", ""));
        assert!(matches!(
            one_module.build(&no_source),
            Err(VbaError::LimitExceeded { .. })
        ));
    }

    #[test]
    fn obfuscation_round_trips_the_unprotected_state_fields() {
        let project_id = VbaProjectId::from_bytes([0xa5; 16]).braced_uppercase();
        for data in [&[0, 0, 0, 0][..], &[0][..], &[0xff][..]] {
            let encrypted = encrypt_project_data(data, &project_id);
            assert_eq!(decrypt_project_data(&encrypted, &project_id), data);
        }
    }

    fn assert_encrypted_property_shape(project: &str, name: &str, hex_length: usize) {
        let prefix = format!("{name}=\"");
        let value = project
            .lines()
            .find_map(|line| line.strip_prefix(&prefix))
            .and_then(|line| line.strip_suffix('"'))
            .unwrap();
        assert_eq!(value.len(), hex_length);
        assert!(value.bytes().all(|byte| byte.is_ascii_hexdigit()));
    }

    fn decrypt_project_data(encrypted: &[u8], project_id: &str) -> Vec<u8> {
        let seed = encrypted[0];
        assert_eq!(encrypted[1] ^ seed, ENCRYPTION_VERSION);
        let project_key = encrypted[2] ^ seed;
        assert_eq!(
            project_key,
            project_id
                .bytes()
                .fold(0u8, |sum, byte| sum.wrapping_add(byte))
        );
        let ignored_length = usize::from((seed & 6) / 2);
        let mut cursor = 3 + ignored_length;
        let mut unencrypted_byte_1 = project_key;
        let mut encrypted_byte_1 = encrypted[2];
        let mut encrypted_byte_2 = encrypted[1];
        for &byte in &encrypted[3..cursor] {
            let plain = byte ^ encrypted_byte_2.wrapping_add(unencrypted_byte_1);
            encrypted_byte_2 = encrypted_byte_1;
            encrypted_byte_1 = byte;
            unencrypted_byte_1 = plain;
        }
        let mut length = [0u8; 4];
        for slot in &mut length {
            let byte = encrypted[cursor];
            cursor += 1;
            let plain = byte ^ encrypted_byte_2.wrapping_add(unencrypted_byte_1);
            *slot = plain;
            encrypted_byte_2 = encrypted_byte_1;
            encrypted_byte_1 = byte;
            unencrypted_byte_1 = plain;
        }
        let length = usize::try_from(u32::from_le_bytes(length)).unwrap();
        let mut output = Vec::with_capacity(length);
        for &byte in &encrypted[cursor..cursor + length] {
            let plain = byte ^ encrypted_byte_2.wrapping_add(unencrypted_byte_1);
            output.push(plain);
            encrypted_byte_2 = encrypted_byte_1;
            encrypted_byte_1 = byte;
            unencrypted_byte_1 = plain;
        }
        output
    }
}
