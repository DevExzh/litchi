//! Typed, inert MS-OVBA project authoring.

use super::dir::{
    Kind as DirKind, Reference, WriteModule, WriteProject, encode_dir, encode_mbcs,
    validate_vba_identifier,
};
use super::project::{DesignerFrame, LicenseInfo};
use super::{Error, Limits, check_limit, codec, invalid};
use litchi_cfb::{OleFile, OleWriter};
use litchi_codepage::Mbcs;
use std::collections::HashSet;
use std::io::{self, Cursor, Seek, SeekFrom, Write};

const VBA_STORAGE_NAME: &str = "VBA";
const DIR_STREAM_NAME: &str = "dir";
const PROJECT_STREAM_NAME: &str = "PROJECT";
const PROJECT_WM_STREAM_NAME: &str = "PROJECTwm";
const PROJECT_LK_STREAM_NAME: &str = "PROJECTlk";
const VERSION_PROJECT_STREAM_NAME: &str = "_VBA_PROJECT";
const VERSION_PROJECT_RESERVED: u16 = 0x61cc;
const VERSION_PROJECT_WRITE_VERSION: u16 = 0xffff;
const DEFAULT_PROJECT_VERSION_MAJOR: u32 = 1;
const DEFAULT_PROJECT_VERSION_MINOR: u16 = 0;
const MAX_CFB_NAME_CODE_UNITS: usize = 31;
const MAX_VBA_IDENTIFIER_CHARACTERS: usize = 31;
const DETERMINISTIC_OBFUSCATION_SEED: u8 = 0;
const ENCRYPTION_VERSION: u8 = 2;
const PROJECT_VERSION_COMPATIBLE_32: &str = "393222000";
const CFB_HEADER_BYTES: usize = 512;

/// Target platform stored in `PROJECTSYSKIND`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Platform {
    /// 16-bit Windows.
    Windows16,
    /// 32-bit Windows.
    Windows32,
    /// Classic Macintosh.
    Macintosh,
    /// 64-bit Windows.
    Windows64,
}

impl Platform {
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
pub enum Kind {
    /// A standard procedural module (`Module=`).
    Standard,
    /// A class module (`Class=`).
    Class,
    /// A designer module with a corresponding VBFrame.
    Designer,
    /// A host document module (`Document=`).
    Document {
        /// Automation server version written after the module name.
        type_library_version: u32,
    },
}

impl Kind {
    const fn directory_kind(self) -> DirKind {
        match self {
            Self::Standard => DirKind::Procedural,
            Self::Class | Self::Designer | Self::Document { .. } => {
                DirKind::DocumentClassOrDesigner
            },
        }
    }
}

/// Canonical project Automation type-library identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Id([u8; 16]);

impl Id {
    /// The all-zero project identifier.
    pub const NIL: Self = Self([0; 16]);

    /// Construct an identifier from bytes in canonical textual GUID order.
    #[must_use]
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    /// Return bytes in canonical textual GUID order.
    #[must_use]
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
#[derive(Debug, PartialEq, Eq)]
pub struct Module {
    name: String,
    stream_name: String,
    kind: Kind,
    source_body: String,
    description: String,
    help_context: u32,
    read_only: bool,
    private: bool,
}

impl Module {
    /// Create a standard procedural module.
    pub fn standard(name: impl Into<String>, source_body: impl Into<String>) -> Self {
        Self::new(name, source_body, Kind::Standard)
    }

    /// Create a class module.
    pub fn class(name: impl Into<String>, source_body: impl Into<String>) -> Self {
        Self::new(name, source_body, Kind::Class)
    }

    /// Create a designer module whose project declaration is BaseClass.
    pub fn designer(name: impl Into<String>, source_body: impl Into<String>) -> Self {
        Self::new(name, source_body, Kind::Designer)
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
            Kind::Document {
                type_library_version,
            },
        )
    }

    fn new(name: impl Into<String>, source_body: impl Into<String>, kind: Kind) -> Self {
        let module_name = name.into();
        Self {
            stream_name: module_name.clone(),
            name: module_name,
            kind,
            source_body: source_body.into(),
            description: String::new(),
            help_context: 0,
            read_only: false,
            private: false,
        }
    }

    /// Override the CFB stream name. By default it equals the module name.
    #[must_use]
    pub fn stream_name(mut self, stream_name: impl Into<String>) -> Self {
        self.stream_name = stream_name.into();
        self
    }

    /// Set the module description.
    #[must_use]
    pub fn description(mut self, description: impl Into<String>) -> Self {
        self.description = description.into();
        self
    }

    /// Set the module Help context identifier.
    #[must_use]
    pub fn help_context(mut self, help_context: u32) -> Self {
        self.help_context = help_context;
        self
    }

    /// Mark the module read-only.
    #[must_use]
    pub fn read_only(mut self, read_only: bool) -> Self {
        self.read_only = read_only;
        self
    }

    /// Mark the module private to its project.
    #[must_use]
    pub fn private(mut self, private: bool) -> Self {
        self.private = private;
        self
    }

    /// Module identifier.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Project-level module category.
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }
}

/// Builder for a cache-free, inert MS-OVBA project.
#[derive(Debug, PartialEq, Eq)]
pub struct Project {
    name: String,
    id: Id,
    platform: Platform,
    page: Mbcs,
    description: String,
    help_context: i32,
    version_major: u32,
    version_minor: u16,
    references: Vec<Reference>,
    license_info: Vec<LicenseInfo>,
    designer_frames: Vec<DesignerFrameInput>,
    modules: Vec<Module>,
}

impl Project {
    /// Create a project with Windows-1252 text and 32-bit Windows metadata.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            id: Id::NIL,
            platform: Platform::Windows32,
            page: Mbcs::WINDOWS_1252,
            description: String::new(),
            help_context: 0,
            version_major: DEFAULT_PROJECT_VERSION_MAJOR,
            version_minor: DEFAULT_PROJECT_VERSION_MINOR,
            references: Vec::new(),
            license_info: Vec::new(),
            designer_frames: Vec::new(),
            modules: Vec::new(),
        }
    }

    /// Set the project Automation type-library identifier.
    #[must_use]
    pub fn id(mut self, project_id: Id) -> Self {
        self.id = project_id;
        self
    }

    /// Set the target platform.
    #[must_use]
    pub fn platform(mut self, platform: Platform) -> Self {
        self.platform = platform;
        self
    }

    /// Set the checked MBCS page used for project text and module source.
    #[must_use]
    pub fn page(mut self, page: Mbcs) -> Self {
        self.page = page;
        self
    }

    /// Validate a raw MBCS identifier and set the project page.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsupportedCodePage`] if `page` is not a checked MBCS
    /// code page known to the encoding layer.
    #[must_use = "project builders must be finished or further configured"]
    pub fn page_id(mut self, page: u16) -> Result<Self, Error> {
        self.page = Mbcs::new(u32::from(page)).ok_or(Error::UnsupportedCodePage(page))?;
        Ok(self)
    }

    /// Set the project description.
    #[must_use]
    pub fn description(mut self, description: impl Into<String>) -> Self {
        self.description = description.into();
        self
    }

    /// Set the project Help context identifier.
    #[must_use]
    pub fn help_context(mut self, help_context: i32) -> Self {
        self.help_context = help_context;
        self
    }

    /// Set the project version mirrored into the `dir` stream.
    #[must_use]
    pub fn version(mut self, major: u32, minor: u16) -> Self {
        self.version_major = major;
        self.version_minor = minor;
        self
    }

    /// Append a module and return the builder.
    #[must_use]
    pub fn module(mut self, module: Module) -> Self {
        self.modules.push(module);
        self
    }

    /// Append an inert external reference record.
    #[must_use]
    pub fn reference(mut self, reference: Reference) -> Self {
        self.references.push(reference);
        self
    }

    /// Append inert ActiveX-control license metadata.
    #[must_use]
    pub fn license_info(mut self, license_info: LicenseInfo) -> Self {
        self.license_info.push(license_info);
        self
    }

    /// Append a validated, inert `VBFrame` stream for a designer module.
    ///
    /// `module_name` must identify a module already added to this builder.
    /// The bytes are retained as source text and are never interpreted as
    /// executable content.
    #[must_use]
    pub fn designer_frame(
        mut self,
        module_name: impl Into<String>,
        raw: impl Into<Vec<u8>>,
    ) -> Self {
        self.designer_frames.push(DesignerFrameInput {
            module_name: module_name.into(),
            raw: raw.into(),
        });
        self
    }

    /// Serialize and validate the project without executing or compiling source.
    ///
    /// # Errors
    ///
    /// Returns an error if project metadata or module source is invalid, a
    /// configured [`Limits`] ceiling is exceeded, or the underlying CFB writer
    /// fails.
    pub fn finish(self, limits: &Limits) -> Result<Payload, Error> {
        check_limit(
            "standalone VBA CFB bytes",
            CFB_HEADER_BYTES,
            limits.max_cfb_bytes,
        )?;
        check_limit("VBA module count", self.modules.len(), limits.max_modules)?;
        check_limit(
            "VBA license count",
            self.license_info.len(),
            limits.max_licenses,
        )?;
        check_limit(
            "VBA designer count",
            self.designer_frames.len(),
            limits.max_modules,
        )?;
        let encoding = self.page;
        validate_project_name(&self.name)?;
        validate_quoted_text(&self.description, "VBA project description")?;
        let encoded_name = encode_mbcs(&self.name, encoding, "PROJECTNAME")?;
        if encoded_name.len() > 128 {
            return Err(invalid("PROJECTNAME exceeds 128 encoded bytes"));
        }
        validate_project_modules(&self.modules)?;

        let mut compressed_modules = Vec::with_capacity(self.modules.len());
        let mut total_source_bytes = 0usize;
        for module in &self.modules {
            let source = module_source(module);
            let encoded_source = encode_mbcs(&source, encoding, "module source")?;
            total_source_bytes = total_source_bytes
                .checked_add(encoded_source.len())
                .ok_or_else(|| invalid("aggregate VBA source size overflow"))?;
            check_limit(
                "aggregate VBA module source bytes",
                total_source_bytes,
                limits.max_total_source_bytes,
            )?;
            let compressed_source = codec::encode(&encoded_source, limits)?;
            compressed_modules.push(compressed_source);
        }

        let directory_modules: Vec<_> = self
            .modules
            .iter()
            .map(|module| WriteModule {
                name: &module.name,
                stream_name: &module.stream_name,
                description: &module.description,
                help_context: module.help_context,
                kind: module.kind.directory_kind(),
                read_only: module.read_only,
                private: module.private,
            })
            .collect();
        let directory = encode_dir(
            &WriteProject {
                system_kind: self.platform.system_kind(),
                page: self.page,
                name: &self.name,
                description: &self.description,
                help_context: u32::from_le_bytes(self.help_context.to_le_bytes()),
                version_major: self.version_major,
                version_minor: self.version_minor,
                references: &self.references,
                modules: &directory_modules,
            },
            limits,
        )?;
        let project_stream = encode_project_stream(&self, encoding, limits)?;
        let project_wm_stream = encode_project_wm_stream(&self.modules, encoding, limits)?;
        let project_lk_stream = encode_project_lk_stream(&self.license_info, limits)?;

        let mut designer_frames = Vec::with_capacity(self.designer_frames.len());
        let mut designer_names = HashSet::with_capacity(self.designer_frames.len());
        for input in &self.designer_frames {
            if !designer_names.insert(input.module_name.to_lowercase()) {
                return Err(invalid(format!(
                    "duplicate VBA designer frame for module {}",
                    input.module_name
                )));
            }
            let module = self
                .modules
                .iter()
                .find(|module| module.name.eq_ignore_ascii_case(&input.module_name))
                .ok_or_else(|| {
                    invalid(format!(
                        "VBA designer frame has no matching module {}",
                        input.module_name
                    ))
                })?;
            if module.kind != Kind::Designer {
                return Err(invalid(format!(
                    "VBA designer frame requires a designer module, found {}",
                    module.name
                )));
            }
            let _ =
                DesignerFrame::from_raw(module.name.clone(), input.raw.clone(), encoding, limits)?;
            designer_frames.push(EncodedDesignerFrame {
                storage_name: module.stream_name.clone(),
                raw: input.raw.clone(),
            });
        }
        for module in &self.modules {
            if module.kind == Kind::Designer
                && !designer_names.contains(&module.name.to_lowercase())
            {
                return Err(invalid(format!(
                    "designer module {} is missing a VBFrame stream",
                    module.name
                )));
            }
        }

        let modules = self
            .modules
            .into_iter()
            .zip(compressed_modules)
            .map(|(module, compressed_source)| EncodedModule {
                stream_name: module.stream_name,
                compressed_source,
            })
            .collect();
        let streams = Streams {
            version_project_stream: version_project_stream(),
            directory_stream: directory,
            project_stream,
            project_wm_stream,
            project_lk_stream,
            designer_frames,
            modules,
        };
        let module_count = streams.modules.len();
        let mut writer = OleWriter::new();
        streams.write_into(&mut writer, &[])?;
        drop(streams);

        let mut output = BoundedCursor::new(limits.max_cfb_bytes);
        if let Err(error) = writer.write_to(&mut output) {
            if let Some(actual) = output.exceeded_actual() {
                return Err(Error::LimitExceeded {
                    limit: "standalone VBA CFB bytes",
                    actual,
                    maximum: limits.max_cfb_bytes,
                });
            }
            return Err(error.into());
        }
        let bytes = output.into_inner();
        check_limit(
            "standalone VBA CFB bytes",
            bytes.len(),
            limits.max_cfb_bytes,
        )?;
        Ok(Payload {
            bytes,
            module_count,
        })
    }
}

struct BoundedCursor {
    inner: Cursor<Vec<u8>>,
    maximum: usize,
    exceeded_actual: Option<usize>,
}

impl BoundedCursor {
    fn new(maximum: usize) -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            maximum,
            exceeded_actual: None,
        }
    }

    fn exceeded_actual(&self) -> Option<usize> {
        self.exceeded_actual
    }

    fn into_inner(self) -> Vec<u8> {
        self.inner.into_inner()
    }

    fn reject_limit(&mut self, input_actual: u64) -> io::Error {
        let actual =
            usize::try_from(input_actual).unwrap_or_else(|_| self.maximum.saturating_add(1));
        self.exceeded_actual = Some(
            self.exceeded_actual
                .map_or(actual, |previous| previous.max(actual)),
        );
        io::Error::other("standalone VBA CFB byte limit exceeded")
    }

    fn maximum_u64(&self) -> u64 {
        u64::try_from(self.maximum).unwrap_or(u64::MAX)
    }
}

impl Write for BoundedCursor {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let Ok(byte_count) = u64::try_from(bytes.len()) else {
            return Err(self.reject_limit(u64::MAX));
        };
        let end = self.inner.position().saturating_add(byte_count);
        if end > self.maximum_u64() {
            return Err(self.reject_limit(end));
        }
        self.inner.write(bytes)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for BoundedCursor {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let previous = self.inner.position();
        let next = self.inner.seek(position)?;
        if next > self.maximum_u64() {
            self.inner.set_position(previous);
            return Err(self.reject_limit(next));
        }
        Ok(next)
    }
}

#[derive(Debug, PartialEq, Eq)]
struct EncodedModule {
    stream_name: String,
    compressed_source: Vec<u8>,
}

#[derive(Debug, PartialEq, Eq)]
struct DesignerFrameInput {
    module_name: String,
    raw: Vec<u8>,
}

#[derive(Debug, PartialEq, Eq)]
struct EncodedDesignerFrame {
    storage_name: String,
    raw: Vec<u8>,
}

#[derive(Debug, PartialEq, Eq)]
struct Streams {
    version_project_stream: Vec<u8>,
    directory_stream: Vec<u8>,
    project_stream: Vec<u8>,
    project_wm_stream: Vec<u8>,
    project_lk_stream: Option<Vec<u8>>,
    designer_frames: Vec<EncodedDesignerFrame>,
    modules: Vec<EncodedModule>,
}

impl Streams {
    fn write_into(&self, writer: &mut OleWriter, project_root_path: &[&str]) -> Result<(), Error> {
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
        if let Some(project_lk_stream) = &self.project_lk_stream {
            project_path.pop();
            project_path.push(PROJECT_LK_STREAM_NAME.to_owned());
            create_stream(writer, &project_path, project_lk_stream)?;
        }
        for designer in &self.designer_frames {
            let mut designer_path: Vec<String> = project_root_path
                .iter()
                .map(|component| (*component).to_owned())
                .collect();
            designer_path.push(designer.storage_name.clone());
            create_storage(writer, &designer_path)?;
            designer_path.push("\u{3}VBFrame".to_owned());
            create_stream(writer, &designer_path, &designer.raw)?;
        }

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
}

/// Owned, validated standalone CFB bytes for one inert MS-OVBA project.
///
/// The private representation is a validation capability: arbitrary bytes can
/// enter only through [`Self::read`], while [`Project::finish`] constructs the
/// same invariant directly.
#[derive(Debug, PartialEq, Eq)]
pub struct Payload {
    bytes: Vec<u8>,
    module_count: usize,
}

impl Payload {
    /// Consume and validate standalone CFB bytes without copying them.
    ///
    /// # Errors
    ///
    /// Returns an error if `bytes` exceeds `limits`, is not a readable CFB
    /// file, or contains a malformed MS-OVBA project.
    pub fn read(bytes: Vec<u8>, limits: &Limits) -> Result<Self, Error> {
        check_limit(
            "standalone VBA CFB bytes",
            bytes.len(),
            limits.max_cfb_bytes,
        )?;
        let mut ole = OleFile::open(Cursor::new(bytes.as_slice()))?;
        let project = crate::project::Project::open(&mut ole, &[], limits)?;
        project.validate_payload(limits)?;
        let _ = project.project_wm(limits)?;
        let module_count = project.modules().len();
        Ok(Self {
            bytes,
            module_count,
        })
    }

    /// Borrow the exact validated standalone bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Move out the exact validated standalone bytes without copying them.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// Number of module streams declared and validated by the project.
    #[must_use]
    pub fn module_count(&self) -> usize {
        self.module_count
    }

    /// Whether the project contains no standard, class, or document modules.
    #[must_use]
    pub fn is_module_free(&self) -> bool {
        self.module_count == 0
    }

    /// Copy all payload streams into an existing CFB writer at a project root.
    ///
    /// # Errors
    ///
    /// Returns an error if the validated payload cannot be re-read, or if the
    /// destination CFB writer rejects a storage or stream.
    pub fn write_into(
        &self,
        writer: &mut OleWriter,
        project_root_path: &[&str],
    ) -> Result<(), Error> {
        let mut source = OleFile::open(Cursor::new(self.bytes.as_slice()))?;
        for source_path in source.list_streams() {
            let is_vba_stream =
                source_path.len() == 2 && source_path[0].eq_ignore_ascii_case(VBA_STORAGE_NAME);
            if is_vba_stream
                && source_path[1]
                    .get(..6)
                    .is_some_and(|prefix| prefix.eq_ignore_ascii_case("__SRP_"))
            {
                continue;
            }
            let source_components: Vec<&str> = source_path.iter().map(String::as_str).collect();
            let data = if is_vba_stream
                && source_path[1].eq_ignore_ascii_case(VERSION_PROJECT_STREAM_NAME)
            {
                version_project_stream()
            } else {
                source.open_stream(&source_components)?
            };
            let mut target_path: Vec<String> = project_root_path
                .iter()
                .map(|component| (*component).to_owned())
                .collect();
            target_path.extend(source_path);
            if target_path.len() > 1 {
                create_storage(writer, &target_path[..target_path.len() - 1])?;
            }
            create_stream(writer, &target_path, &data)?;
        }
        Ok(())
    }
}

fn create_storage(writer: &mut OleWriter, path: &[String]) -> Result<(), Error> {
    let components: Vec<&str> = path.iter().map(String::as_str).collect();
    writer.create_storage(&components)?;
    Ok(())
}

fn create_stream(writer: &mut OleWriter, path: &[String], data: &[u8]) -> Result<(), Error> {
    let components: Vec<&str> = path.iter().map(String::as_str).collect();
    writer.create_stream(&components, data)?;
    Ok(())
}

fn validate_project_name(name: &str) -> Result<(), Error> {
    if name.is_empty() {
        return Err(invalid("VBA project name must not be empty"));
    }
    if name.chars().count() > 128 {
        return Err(invalid("VBA project name exceeds 128 characters"));
    }
    if name
        .chars()
        .any(|character| !valid_quoted_character(character))
    {
        return Err(invalid(
            "VBA project name contains a character that cannot appear in QUOTEDCHAR",
        ));
    }
    validate_vba_identifier(name, "PROJECTNAME")?;
    Ok(())
}

fn validate_project_modules(modules: &[Module]) -> Result<(), Error> {
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

fn validate_module_identifier(name: &str) -> Result<(), Error> {
    if name.chars().count() > MAX_VBA_IDENTIFIER_CHARACTERS {
        return Err(invalid(format!(
            "VBA module name exceeds {MAX_VBA_IDENTIFIER_CHARACTERS} characters"
        )));
    }
    if name.is_empty() {
        return Err(invalid("VBA module name must not be empty"));
    }
    if name
        .chars()
        .any(|character| matches!(character, '\0' | '\r' | '\n'))
    {
        return Err(invalid(
            "VBA module name contains a forbidden line or null character",
        ));
    }
    Ok(())
}

fn validate_stream_name(name: &str) -> Result<(), Error> {
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

fn validate_quoted_text(value: &str, field: &'static str) -> Result<(), Error> {
    if value
        .chars()
        .any(|character| !valid_quoted_character(character))
    {
        return Err(invalid(format!(
            "{field} contains a character that cannot appear in QUOTEDCHAR"
        )));
    }
    Ok(())
}

fn module_source(module: &Module) -> String {
    let mut source = String::with_capacity(module.name.len() + module.source_body.len() + 32);
    source.push_str("Attribute VB_Name = \"");
    source.push_str(&quote_project_value(&module.name));
    source.push_str("\"\r\n");
    source.push_str(&module.source_body);
    source
}

fn encode_project_stream(
    project: &Project,
    encoding: Mbcs,
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    let project_id = project.id.braced_uppercase();
    let mut text = String::new();
    text.push_str("ID=\"");
    text.push_str(&project_id);
    text.push_str("\"\r\n");
    for module in &project.modules {
        match module.kind {
            Kind::Standard => text.push_str("Module="),
            Kind::Class => text.push_str("Class="),
            Kind::Designer => text.push_str("BaseClass="),
            Kind::Document {
                type_library_version,
            } => {
                if type_library_version > i32::MAX as u32 {
                    return Err(invalid(
                        "VBA document type-library version is not a signed HEXINT32",
                    ));
                }
                text.push_str("Document=");
                text.push_str(&module.name);
                text.push_str("/&H");
                push_upper_hex(&mut text, &type_library_version.to_be_bytes());
                text.push_str("\r\n");
                continue;
            },
        }
        text.push_str(&module.name);
        text.push_str("\r\n");
    }
    text.push_str("Name=\"");
    text.push_str(&quote_project_value(&project.name));
    text.push_str("\"\r\n");
    text.push_str("HelpContextID=\"");
    text.push_str(&project.help_context.to_string());
    text.push_str("\"\r\n");
    if !project.description.is_empty() {
        text.push_str("Description=\"");
        text.push_str(&quote_project_value(&project.description));
        text.push_str("\"\r\n");
    }
    text.push_str("VersionCompatible32=\"");
    text.push_str(PROJECT_VERSION_COMPATIBLE_32);
    text.push_str("\"\r\n");

    let protection = encrypt_project_data(&[0; 4], &project_id)?;
    let password = encrypt_project_data(&[0], &project_id)?;
    let visibility = encrypt_project_data(&[0xff], &project_id)?;
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

fn valid_quoted_character(character: char) -> bool {
    character == '\t' || !character.is_control() || ('\u{7f}'..='\u{ff}').contains(&character)
}

fn quote_project_value(value: &str) -> String {
    value.replace('"', "\"\"")
}

fn encode_project_wm_stream(
    modules: &[Module],
    encoding: Mbcs,
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
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

fn encode_project_lk_stream(
    entries: &[LicenseInfo],
    limits: &Limits,
) -> Result<Option<Vec<u8>>, Error> {
    if entries.is_empty() {
        return Ok(None);
    }
    check_limit("VBA license count", entries.len(), limits.max_licenses)?;
    let count =
        u32::try_from(entries.len()).map_err(|_| invalid("PROJECTlk license count exceeds u32"))?;
    let mut output = Vec::new();
    output.extend_from_slice(&1u16.to_le_bytes());
    output.extend_from_slice(&count.to_le_bytes());
    for entry in entries {
        check_limit(
            "VBA license key bytes",
            entry.license_key().len(),
            limits.max_string_bytes,
        )?;
        let key_length = u32::try_from(entry.license_key().len())
            .map_err(|_| invalid("PROJECTlk license key length exceeds u32"))?;
        let record_length = 16usize
            .checked_add(4)
            .and_then(|length| length.checked_add(entry.license_key().len()))
            .and_then(|length| length.checked_add(4))
            .ok_or_else(|| invalid("PROJECTlk stream size overflow"))?;
        let next_length = output
            .len()
            .checked_add(record_length)
            .ok_or_else(|| invalid("PROJECTlk stream size overflow"))?;
        check_limit(
            "PROJECTlk stream bytes",
            next_length,
            limits.max_decompressed_stream_bytes,
        )?;
        output.extend_from_slice(&entry.class_id());
        output.extend_from_slice(&key_length.to_le_bytes());
        output.extend_from_slice(entry.license_key());
        output.extend_from_slice(&entry.license_required_raw().to_le_bytes());
    }
    Ok(Some(output))
}

fn version_project_stream() -> Vec<u8> {
    let mut output = Vec::with_capacity(7);
    output.extend_from_slice(&VERSION_PROJECT_RESERVED.to_le_bytes());
    output.extend_from_slice(&VERSION_PROJECT_WRITE_VERSION.to_le_bytes());
    output.push(0);
    output.extend_from_slice(&0u16.to_le_bytes());
    output
}

fn encrypt_project_data(data: &[u8], project_id: &str) -> Result<Vec<u8>, Error> {
    let seed = DETERMINISTIC_OBFUSCATION_SEED;
    let project_key = project_id.bytes().fold(0u8, u8::wrapping_add);
    let version_encrypted = seed ^ ENCRYPTION_VERSION;
    let project_key_encrypted = seed ^ project_key;
    let mut output = Vec::with_capacity(7 + data.len());
    output.extend([seed, version_encrypted, project_key_encrypted]);

    let mut unencrypted_byte_1 = project_key;
    let mut encrypted_byte_1 = project_key_encrypted;
    let mut encrypted_byte_2 = version_encrypted;
    let Ok(data_length) = u32::try_from(data.len()) else {
        return Err(invalid("PROJECT encryption input exceeds u32"));
    };
    for byte in data_length
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
    Ok(output)
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
    #![allow(
        clippy::unwrap_used,
        reason = "test fixtures and assertions panic intentionally on failure"
    )]

    use super::*;
    use litchi_cfb::OleFile;

    fn sample_builder() -> Project {
        Project::new("SampleProject")
            .id(Id::from_bytes([
                0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc, 0xde, 0xf0, 0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc,
                0xde, 0xf0,
            ]))
            .description("Inert test project")
            .version(7, 3)
            .module(
                Module::standard("Module1", "Sub Main()\r\nEnd Sub\r\n")
                    .read_only(true)
                    .private(true),
            )
            .module(Module::designer("Class1", "Private value As Long\r\n"))
            .module(Module::document(
                "ThisDocument",
                0x0001_0000,
                "Private Sub Document_Open()\r\nEnd Sub\r\n",
            ))
            .license_info(LicenseInfo::new([0x11; 16], vec![1, 2, 3], true))
            .designer_frame(
                "Class1",
                concat!(
                    "VERSION 5.00\r\n",
                    "Begin {00000000-0000-0000-0000-000000000003} Class1\r\n",
                    "Caption=\"Designer\"\r\n",
                    "Enabled=-1\r\n",
                    "End\r\n",
                ),
            )
    }

    #[test]
    fn writes_complete_cache_free_project_and_reopens_every_module() {
        let limits = Limits::default();
        let binary = sample_builder().finish(&limits).unwrap();
        assert_eq!(binary.module_count(), 3);
        assert!(!binary.is_module_free());
        assert_eq!(
            binary,
            sample_builder().finish(&limits).unwrap(),
            "authoring must be deterministic"
        );

        let bytes = binary.into_bytes();
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

        let project = crate::project::Project::open(&mut ole, &[], &limits).unwrap();
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
            DirKind::DocumentClassOrDesigner
        );
        assert_eq!(project.modules()[2].name(), "ThisDocument");
        let project_wm = project.project_wm(&limits).unwrap().unwrap();
        assert_eq!(project_wm.len(), 3);
        assert_eq!(project_wm[0].mbcs_name(), "Module1");
        assert_eq!(project_wm[0].unicode_name(), "Module1");
        let licenses = project.license_info().unwrap();
        assert_eq!(licenses.len(), 1);
        assert_eq!(licenses[0].class_id(), [0x11; 16]);
        assert_eq!(licenses[0].license_key(), [1, 2, 3]);
        assert!(licenses[0].license_required());
        assert_eq!(project.designer_frames().len(), 1);
        assert_eq!(project.designer_frames()[0].module_name(), "Class1");
        assert_eq!(
            project.designer_frames()[0].class_id(),
            "{00000000-0000-0000-0000-000000000003}"
        );
        assert_eq!(
            project.designer_frames()[0].properties()[0].name(),
            "Caption"
        );

        let properties = project.project_properties().text();
        let project_text = project.project_text(&limits).unwrap();
        assert_eq!(
            project.encode_project_text(project_text, &limits).unwrap(),
            project.project_properties().raw(),
            "an unchanged PROJECT text view must preserve the exact source"
        );
        assert_eq!(project_text.project_name(), Some("SampleProject"));
        assert!(
            project_text
                .records()
                .iter()
                .any(|record| matches!(record, crate::project::ProjectTextRecord::Module { .. }))
        );
        assert!(properties.contains("Module=Module1\r\n"));
        assert!(properties.contains("BaseClass=Class1\r\n"));
        assert!(properties.contains("Document=ThisDocument/&H00010000\r\n"));
        assert!(properties.contains("Description=\"Inert test project\"\r\n"));
        assert!(properties.contains("\r\n[Host Extender Info]\r\n"));
        assert_encrypted_property_shape(properties, "CMG", 22);
        assert_encrypted_property_shape(properties, "DPB", 16);
        assert_encrypted_property_shape(properties, "GC", 16);
    }

    #[test]
    fn writes_nested_legacy_project_and_non_ascii_codepage_text() {
        let builder = Project::new("日本語")
            .page(Mbcs::SHIFT_JIS)
            .platform(Platform::Windows64)
            .module(Module::standard(
                "標準",
                "Sub 挨拶()\r\nMsgBox \"こんにちは\"\r\nEnd Sub\r\n",
            ));
        let binary = builder.finish(&Limits::default()).unwrap();

        let mut writer = OleWriter::new();
        binary
            .write_into(&mut writer, &["_VBA_PROJECT_CUR"])
            .unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        bytes.set_position(0);
        let mut ole = OleFile::open(bytes).unwrap();
        let project =
            crate::project::Project::open(&mut ole, &["_VBA_PROJECT_CUR"], &Limits::default())
                .unwrap();
        assert_eq!(project.name(), "日本語");
        assert_eq!(project.page(), Mbcs::SHIFT_JIS);
        assert!(project.modules()[0].source().text().contains("こんにちは"));
        assert!(!project.modules()[0].source().had_decode_errors());
    }

    #[test]
    fn project_wm_contains_ordered_mbcs_and_unicode_name_pairs() {
        let limits = Limits::default();
        let builder = Project::new("Map")
            .page(Mbcs::SHIFT_JIS)
            .module(Module::standard("標準", ""))
            .module(Module::class("Class1", ""));
        let binary = builder.finish(&limits).unwrap();
        let encoding = Mbcs::SHIFT_JIS;
        let mut expected = Vec::new();
        for name in ["標準", "Class1"] {
            expected.extend_from_slice(&encode_mbcs(name, encoding, "test").unwrap());
            expected.push(0);
            expected.extend(name.encode_utf16().flat_map(u16::to_le_bytes));
            expected.extend_from_slice(&0u16.to_le_bytes());
        }
        expected.extend_from_slice(&0u16.to_le_bytes());
        let mut ole = OleFile::open(Cursor::new(binary.bytes())).unwrap();
        assert_eq!(
            ole.open_stream(&[PROJECT_WM_STREAM_NAME]).unwrap(),
            expected
        );
        let project = crate::project::Project::open(&mut ole, &[], &limits).unwrap();
        let maps = project.project_wm(&limits).unwrap().unwrap().to_vec();
        assert_eq!(project.encode_project_wm(&maps, &limits).unwrap(), expected);
        let mut edited = maps;
        edited[0] = crate::dir::NameMap::new("Class1", "標準");
        assert!(matches!(
            project.encode_project_wm(&edited, &limits),
            Err(Error::InvalidData(message)) if message.contains("MBCS and Unicode")
        ));
        assert!(matches!(
            project.encode_project_wm(&edited[..1], &limits),
            Err(Error::InvalidData(message)) if message.contains("map count")
        ));
    }

    #[test]
    fn references_are_typed_and_reopenable_without_resolution() {
        let original = crate::dir::ControlReference::new(
            "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##",
            Some(crate::dir::ExtendedReference::new(
                Some("MSForms".to_owned()),
                "*\\G{896C2D83-5466-46ED-8FAE-4C3E4F85E710}#2.0#0##",
                [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16],
                7,
            )),
        );
        let builder = Project::new("Refs")
            .reference(Reference::registered(
                "stdole",
                "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0##OLE#Automation",
            ))
            .reference(Reference::project(
                "OtherProject",
                "*\\Cabsolute.xls",
                "*\\Crelative.xls",
                3,
                2,
            ))
            .reference(Reference::control(
                "MSForms",
                "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##",
                original.extended().cloned(),
            ))
            .reference(Reference::original(
                Some("MSFormsOriginal".to_owned()),
                "*\\G{11111111-1111-1111-1111-111111111111}#2.0#0##",
                original,
            ))
            .reference(Reference::opaque(0x7f00, [9, 8, 7, 6]))
            .module(Module::standard("Module1", ""));
        let binary = builder.finish(&Limits::default()).unwrap();
        let mut ole = OleFile::open(Cursor::new(binary.bytes())).unwrap();
        let project = crate::project::Project::open(&mut ole, &[], &Limits::default()).unwrap();
        let references = project.references();
        assert_eq!(references.len(), 5);
        assert!(matches!(
            references[0].kind(),
            crate::dir::ReferenceKind::Registered { libid } if libid.contains("00020430")
        ));
        assert!(matches!(
            references[1].kind(),
            crate::dir::ReferenceKind::Project {
                major_version: 3,
                minor_version: 2,
                ..
            }
        ));
        assert!(matches!(
            references[2].kind(),
            crate::dir::ReferenceKind::Control {
                extended: Some(_),
                ..
            }
        ));
        assert!(matches!(
            references[3].kind(),
            crate::dir::ReferenceKind::Original { .. }
        ));
        assert!(matches!(
            references[4].kind(),
            crate::dir::ReferenceKind::Opaque { id: 0x7f00, payload } if payload == &[9, 8, 7, 6]
        ));
        let source_bound_name = references[0]
            .clone()
            .edit_name_source_bound(
                Some("stdole-edited".to_owned()),
                Mbcs::WINDOWS_1252,
                &Limits::default(),
            )
            .unwrap();
        assert_eq!(source_bound_name.name(), Some("stdole-edited"));
        assert!(source_bound_name.raw().is_some());
        assert!(matches!(
            source_bound_name.kind(),
            crate::dir::ReferenceKind::Registered { .. }
        ));
        assert_eq!(
            project
                .encode_dir_with_references(references, &Limits::default())
                .unwrap(),
            project.dir_raw(),
            "an unchanged reference array must preserve the compressed dir source"
        );
        let mut edited_references = references.to_vec();
        edited_references[0] = source_bound_name;
        let edited_dir = project
            .encode_dir_with_references(&edited_references, &Limits::default())
            .unwrap();
        let edited_directory = crate::dir::Dir::parse(&edited_dir, &Limits::default()).unwrap();
        assert_eq!(
            edited_directory.references()[0].name(),
            Some("stdole-edited")
        );
        assert_eq!(
            edited_directory.modules()[0].name(),
            project.modules()[0].name(),
            "reference edits must retain the module records"
        );
    }

    #[test]
    fn rejects_malformed_reference_abnf_known_framing_ids_and_limits() {
        let malformed_libid = Project::new("BadRef").reference(Reference::registered(
            "stdole",
            "*\\G{00000000-0000-0000-0000-000000000000}#1#0##",
        ));
        assert!(matches!(
            malformed_libid.finish(&Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("LibidReference")
        ));

        let malformed_project = Project::new("BadRef").reference(Reference::project(
            "Other",
            "*\\Eproject.xls",
            "*\\Cproject.xls",
            1,
            0,
        ));
        assert!(matches!(
            malformed_project.finish(&Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("ProjectReference")
        ));

        for reference in [
            Reference::project("", "*\\Aproject.xls", "*\\Cproject.xls", 1, 0),
            Reference::registered("", "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##"),
        ] {
            assert!(matches!(
                Project::new("BadRef")
                    .reference(reference)
                    .finish(&Limits::default()),
                Err(Error::InvalidData(message)) if message.contains("REFERENCENAME")
            ));
        }

        let hyphenated = Project::new("BadRef")
            .reference(Reference::registered(
                "stdole-edited",
                "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##",
            ))
            .module(Module::standard("Module1", ""));
        assert!(matches!(
            hyphenated.finish(&Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("REFERENCENAME")
        ));

        let extended = crate::dir::ExtendedReference::new(
            Some("stdole-edited".to_owned()),
            "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##",
            [0; 16],
            0,
        );
        let control_hyphenated = Project::new("BadRef")
            .reference(Reference::control(
                "MSForms",
                "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##",
                Some(extended),
            ))
            .module(Module::standard("Module1", ""));
        assert!(matches!(
            control_hyphenated.finish(&Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("REFERENCENAME")
        ));

        let opaque_framing = Project::new("BadRef").reference(Reference::opaque(0x000f, []));
        assert!(matches!(
            opaque_framing.finish(&Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("known directory record")
        ));

        let long = "a".repeat(65);
        let limits = Limits {
            max_string_bytes: 16,
            ..Limits::default()
        };
        let bounded = Project::new("BadRef").reference(Reference::registered("stdole", long));
        assert!(matches!(
            bounded.finish(&limits),
            Err(Error::LimitExceeded {
                limit: "VBA input string bytes",
                ..
            })
        ));
    }

    #[test]
    fn payload_validation_and_extraction_preserve_the_input_allocation() {
        let bytes = sample_builder()
            .finish(&Limits::default())
            .unwrap()
            .into_bytes();
        let expected = bytes.clone();
        let pointer = bytes.as_ptr();
        let capacity = bytes.capacity();

        let payload = Payload::read(bytes, &Limits::default()).unwrap();
        assert_eq!(payload.bytes().as_ptr(), pointer);
        assert_eq!(payload.bytes.capacity(), capacity);
        assert_eq!(payload.bytes(), expected);

        let returned_bytes = payload.into_bytes();
        assert_eq!(returned_bytes.as_ptr(), pointer);
        assert_eq!(returned_bytes.capacity(), capacity);
        assert_eq!(returned_bytes, expected);
    }

    fn rewrite_root_stream(payload: &[u8], stream_name: &str, replacement: &[u8]) -> Vec<u8> {
        rewrite_stream_path(payload, &[stream_name], replacement)
    }

    fn rewrite_stream_path(
        payload: &[u8],
        replacement_path: &[&str],
        replacement: &[u8],
    ) -> Vec<u8> {
        let mut source = OleFile::open(Cursor::new(payload)).unwrap();
        let mut writer = OleWriter::new();
        for path in source.list_streams() {
            let components: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = if path.len() == replacement_path.len()
                && path
                    .iter()
                    .zip(replacement_path)
                    .all(|(actual, expected)| actual.eq_ignore_ascii_case(expected))
            {
                replacement.to_vec()
            } else {
                source.open_stream(&components).unwrap()
            };
            if components.len() > 1 {
                writer
                    .create_storage(&components[..components.len() - 1])
                    .unwrap();
            }
            writer.create_stream(&components, &data).unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn remove_stream_with_suffix(payload: &[u8], suffix: &str) -> Vec<u8> {
        let mut source = OleFile::open(Cursor::new(payload)).unwrap();
        let mut writer = OleWriter::new();
        for path in source.list_streams() {
            if path
                .last()
                .is_some_and(|name| name.eq_ignore_ascii_case(suffix))
            {
                continue;
            }
            let components: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = source.open_stream(&components).unwrap();
            if components.len() > 1 {
                writer
                    .create_storage(&components[..components.len() - 1])
                    .unwrap();
            }
            writer.create_stream(&components, &data).unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn owned_payload_admission_eagerly_validates_project_and_project_wm() {
        let payload = sample_builder().finish(&Limits::default()).unwrap();
        let malformed_project = rewrite_root_stream(payload.bytes(), "PROJECT", b"Name=\"P\r\n");
        assert!(matches!(
            Payload::read(malformed_project, &Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("PROJECT Name")
        ));

        let malformed_wm = rewrite_root_stream(payload.bytes(), "PROJECTwm", &[1]);
        assert!(matches!(
            Payload::read(malformed_wm, &Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("PROJECTwm")
        ));

        let designer = Project::new("Designer")
            .module(Module::designer("Form1", ""))
            .designer_frame(
                "Form1",
                b"VERSION 5.00\r\nBegin {00000000-0000-0000-0000-000000000000} Form1\r\nEnd\r\n"
                    .to_vec(),
            )
            .finish(&Limits::default())
            .unwrap();
        let without_frame = remove_stream_with_suffix(designer.bytes(), "\u{3}VBFrame");
        assert!(matches!(
            Payload::read(without_frame, &Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("BaseClass Form1")
        ));
    }

    #[test]
    fn owned_payload_admission_closes_project_and_dir_structure() {
        let limits = Limits::default();
        let payload = Project::new("P")
            .module(Module::standard("M", ""))
            .finish(&limits)
            .unwrap();
        for text in [
            "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\nModule=M\r\nName=\"Q\"\r\nHelpContextID=\"0\"\r\nVersionCompatible32=\"393222000\"\r\nCMG=\"0000000000000000000000\"\r\nDPB=\"0000000000000000\"\r\nGC=\"0000000000000000\"\r\n\r\n[Host Extender Info]\r\n",
            "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\nName=\"P\"\r\nHelpContextID=\"0\"\r\nVersionCompatible32=\"393222000\"\r\nCMG=\"0000000000000000000000\"\r\nDPB=\"0000000000000000\"\r\nGC=\"0000000000000000\"\r\n\r\n[Host Extender Info]\r\n",
            "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\nModule=M\r\nModule=Ghost\r\nName=\"P\"\r\nHelpContextID=\"0\"\r\nVersionCompatible32=\"393222000\"\r\nCMG=\"0000000000000000000000\"\r\nDPB=\"0000000000000000\"\r\nGC=\"0000000000000000\"\r\n\r\n[Host Extender Info]\r\n",
            "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\nModule=M\r\nName=\"P\"\r\nHelpContextID=\"0\"\r\nVersionCompatible32=\"393222000\"\r\nCMG=\"0000000000000000000000\"\r\nDPB=\"0000000000000000\"\r\nGC=\"0000000000000000\"\r\n\r\n[Host Extender Info]\r\n\r\n[Workspace]\r\nGhost=0, 0, 0, 0, C\r\n",
            "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\nModule=M\r\nName=\"P\"\r\nHelpContextID=\"0\"\r\nVersionCompatible32=\"393222000\"\r\nCMG=\"0000000000000000000000\"\r\nDPB=\"0000000000000000\"\r\nGC=\"0000000000000000\"\r\n\r\n[Host Extender Info]\r\n[Workspace]\r\nM=0, 0, 0, 0, C\r\n",
            "Name=\"P\"\r\n",
        ] {
            assert!(
                Payload::read(
                    rewrite_root_stream(payload.bytes(), "PROJECT", text.as_bytes()),
                    &limits,
                )
                .is_err(),
                "malformed PROJECT text must be rejected during owned admission"
            );
        }
        let valid_workspace = concat!(
            "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n",
            "Module=M\r\n",
            "Name=\"P\"\r\n",
            "HelpContextID=\"0\"\r\n",
            "VersionCompatible32=\"393222000\"\r\n",
            "CMG=\"0000000000000000000000\"\r\n",
            "DPB=\"0000000000000000\"\r\n",
            "GC=\"0000000000000000\"\r\n",
            "\r\n[Host Extender Info]\r\n",
            "\r\n[Workspace]\r\n",
            "M=0, 0, 0, 0, C\r\n",
        );
        assert!(
            Payload::read(
                rewrite_root_stream(payload.bytes(), "PROJECT", valid_workspace.as_bytes()),
                &limits,
            )
            .is_ok(),
            "a matching workspace window with the required separator is valid"
        );

        let mut source = OleFile::open(Cursor::new(payload.bytes())).unwrap();
        let compressed_dir = source.open_stream(&["VBA", DIR_STREAM_NAME]).unwrap();
        let mut dir = codec::decode(&compressed_dir, &limits).unwrap();
        dir.extend_from_slice(&[0xaa, 0xbb]);
        let compressed = codec::encode(&dir, &limits).unwrap();
        let malformed =
            rewrite_stream_path(payload.bytes(), &["VBA", DIR_STREAM_NAME], &compressed);
        assert!(matches!(
            Payload::read(malformed, &limits),
            Err(Error::InvalidData(message)) if message.contains("trailing bytes")
        ));
    }

    #[test]
    fn output_limit_stops_before_full_payload_allocation() {
        for maximum in [0, CFB_HEADER_BYTES - 1] {
            let limits = Limits {
                max_cfb_bytes: maximum,
                max_decompressed_stream_bytes: 0,
                max_total_source_bytes: 0,
                ..Limits::default()
            };
            assert!(matches!(
                sample_builder().finish(&limits),
                Err(Error::LimitExceeded {
                    limit: "standalone VBA CFB bytes",
                    actual: CFB_HEADER_BYTES,
                    maximum: observed,
                }) if observed == maximum
            ));
        }

        let full_size = sample_builder()
            .finish(&Limits::default())
            .unwrap()
            .bytes()
            .len();
        let maximum = full_size - 1;
        let limits = Limits {
            max_cfb_bytes: maximum,
            ..Limits::default()
        };
        assert!(matches!(
            sample_builder().finish(&limits),
            Err(Error::LimitExceeded {
                limit: "standalone VBA CFB bytes",
                actual,
                maximum: observed,
            }) if actual > observed && observed == maximum
        ));

        let mut output = BoundedCursor::new(8);
        output.write_all(&[0; 8]).unwrap();
        let capacity = output.inner.get_ref().capacity();
        assert!(output.write_all(&[1]).is_err());
        assert_eq!(output.inner.get_ref().len(), 8);
        assert_eq!(output.inner.get_ref().capacity(), capacity);
        assert_eq!(output.exceeded_actual(), Some(9));
    }

    #[test]
    fn rejects_invalid_names_codepages_text_and_resource_limits() {
        let limits = Limits::default();
        for name in ["1Project", "Project Name", "End", "P!"] {
            assert!(
                Project::new(name).finish(&limits).is_err(),
                "fresh PROJECTNAME {name:?} must satisfy VbaIdentifier"
            );
        }
        let duplicate = || {
            Project::new("Project")
                .module(Module::standard("Module1", ""))
                .module(Module::class("module1", ""))
        };
        assert!(duplicate().finish(&limits).is_err());

        let reserved =
            Project::new("Project").module(Module::standard("Module1", "").stream_name("dir"));
        assert!(reserved.finish(&limits).is_err());
        let cache_name =
            Project::new("Project").module(Module::standard("Module1", "").stream_name("__SRP_0"));
        assert!(cache_name.finish(&limits).is_err());

        let invalid_identifier =
            Project::new("Project").module(Module::standard("bad\0module", ""));
        assert!(invalid_identifier.finish(&limits).is_err());
        let arbitrary_identifier = Project::new("Project")
            .description("quoted \"description\"\twith DEL \u{7f}")
            .module(Module::standard("1 Module", ""));
        let arbitrary_payload = arbitrary_identifier.finish(&limits).unwrap();
        let mut arbitrary_ole = OleFile::open(Cursor::new(arbitrary_payload.bytes())).unwrap();
        let arbitrary_project =
            crate::project::Project::open(&mut arbitrary_ole, &[], &limits).unwrap();
        assert_eq!(arbitrary_project.modules()[0].name(), "1 Module");
        assert!(
            arbitrary_project
                .project_properties()
                .text()
                .contains("Description=\"quoted \"\"description\"\"\twith DEL \u{7f}\"")
        );
        let long_identifier = Project::new("Project").module(Module::standard("a".repeat(32), ""));
        assert!(matches!(
            long_identifier.finish(&limits),
            Err(Error::InvalidData(message)) if message.contains("31 characters")
        ));

        let quoted = Project::new("Project")
            .description("valid \"description")
            .finish(&limits)
            .unwrap();
        let mut quoted_ole = OleFile::open(Cursor::new(quoted.bytes())).unwrap();
        let quoted_project = crate::project::Project::open(&mut quoted_ole, &[], &limits).unwrap();
        assert!(
            quoted_project
                .project_properties()
                .text()
                .contains("Description=\"valid \"\"description\"")
        );
        let unrepresentable =
            Project::new("Project").module(Module::standard("Module1", "MsgBox \"🙂\""));
        assert!(unrepresentable.finish(&limits).is_err());
        assert!(matches!(
            Project::new("Project").page_id(42),
            Err(Error::UnsupportedCodePage(42))
        ));

        let no_modules = Limits {
            max_modules: 0,
            ..limits
        };
        assert!(duplicate().finish(&no_modules).is_err());
        let no_source = Limits {
            max_total_source_bytes: 0,
            ..limits
        };
        let one_module = Project::new("Project").module(Module::standard("Module1", ""));
        assert!(matches!(
            one_module.finish(&no_source),
            Err(Error::LimitExceeded { .. })
        ));

        let invalid_reference_names = ["Bad Name", "1Bad", "End", "P!"];
        for name in invalid_reference_names {
            assert!(
                Project::new("Project")
                    .reference(Reference::project(
                        name,
                        "*\\Aproject.xls",
                        "*\\Cproject.xls",
                        1,
                        0,
                    ))
                    .finish(&limits)
                    .is_err(),
                "fresh project reference name {name:?} must satisfy RefProjectName"
            );
            assert!(
                Project::new("Project")
                    .reference(Reference::registered(
                        name,
                        "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##",
                    ))
                    .finish(&limits)
                    .is_err(),
                "fresh library reference name {name:?} must satisfy RefLibraryName"
            );
        }
    }

    #[test]
    fn designer_frames_require_baseclass_modules_and_round_trip_as_baseclass() {
        let frame = concat!(
            "VERSION 5.00\r\n",
            "Begin {00000000-0000-0000-0000-000000000000} Form1\r\n",
            "End\r\n",
        );
        let payload = Project::new("Designer")
            .module(Module::designer("Form1", ""))
            .designer_frame("Form1", frame.as_bytes().to_vec())
            .finish(&Limits::default())
            .unwrap();
        let mut ole = OleFile::open(Cursor::new(payload.bytes())).unwrap();
        let project = crate::project::Project::open(&mut ole, &[], &Limits::default()).unwrap();
        assert_eq!(project.designer_frames().len(), 1);
        assert!(
            project
                .project_properties()
                .text()
                .contains("BaseClass=Form1\r\n")
        );

        let class_with_frame = Project::new("Designer")
            .module(Module::class("Form1", ""))
            .designer_frame("Form1", frame.as_bytes().to_vec());
        assert!(matches!(
            class_with_frame.finish(&Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("requires a designer module")
        ));

        assert!(matches!(
            Project::new("Designer")
                .module(Module::designer("Form1", ""))
                .finish(&Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("missing a VBFrame")
        ));
    }

    #[test]
    fn obfuscation_round_trips_the_unprotected_state_fields() {
        let project_id = Id::from_bytes([0xa5; 16]).braced_uppercase();
        for data in [&[0, 0, 0, 0][..], &[0][..], &[0xff][..]] {
            let encrypted = encrypt_project_data(data, &project_id).unwrap();
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
        assert_eq!(project_key, project_id.bytes().fold(0u8, u8::wrapping_add));
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
        let data_length = usize::try_from(u32::from_le_bytes(length)).unwrap();
        let mut output = Vec::with_capacity(data_length);
        for &byte in &encrypted[cursor..cursor + data_length] {
            let plain = byte ^ encrypted_byte_2.wrapping_add(unencrypted_byte_1);
            output.push(plain);
            encrypted_byte_2 = encrypted_byte_1;
            encrypted_byte_1 = byte;
            unencrypted_byte_1 = plain;
        }
        output
    }
}
