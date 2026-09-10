//! CFB-backed loading of complete inert MS-OVBA projects.

use super::dir::{
    Dir, Kind, NameMap, Reference, decode_utf16, encode_dir_with_references,
    validate_vba_identifier,
};
use super::{Error, Limits, check_limit, codec, invalid};
use litchi_cfb::{OleError, OleFile};
use litchi_codepage::Mbcs;
use std::fmt::Write as FmtWrite;
use std::io::{Cursor, Read, Seek};
use std::sync::OnceLock;

const VBA_STORAGE_NAME: &str = "VBA";
const DIR_STREAM_NAME: &str = "dir";
const PROJECT_STREAM_NAME: &str = "PROJECT";
const VERSION_PROJECT_STREAM_NAME: &str = "_VBA_PROJECT";
const VERSION_PROJECT_HEADER_BYTES: usize = 7;
const PROJECT_WM_STREAM_NAME: &str = "PROJECTwm";
const PROJECT_LK_STREAM_NAME: &str = "PROJECTlk";
const VBFRAME_STREAM_PREFIX: &str = "VBFrame";

/// One ActiveX-control license entry from the optional `PROJECTlk` stream.
///
/// The license key is retained as opaque bytes.  litchi-vba does not activate
/// controls or interpret the key, and the raw Boolean value is retained so a
/// source-preserving caller can distinguish non-canonical producer values.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LicenseInfo {
    class_id: [u8; 16],
    license_key: Vec<u8>,
    license_required: u32,
}

impl LicenseInfo {
    /// Construct a license entry from its CLSID, key bytes, and Boolean value.
    #[must_use]
    pub fn new(class_id: [u8; 16], license_key: Vec<u8>, license_required: bool) -> Self {
        Self {
            class_id,
            license_key,
            license_required: u32::from(license_required),
        }
    }

    /// Construct an entry while retaining the exact on-wire Boolean value.
    #[must_use]
    pub fn from_raw(class_id: [u8; 16], license_key: Vec<u8>, license_required: u32) -> Self {
        Self {
            class_id,
            license_key,
            license_required,
        }
    }

    /// ActiveX control class identifier bytes in the stream's GUID order.
    #[must_use]
    pub fn class_id(&self) -> [u8; 16] {
        self.class_id
    }

    /// Opaque license-key bytes.
    #[must_use]
    pub fn license_key(&self) -> &[u8] {
        &self.license_key
    }

    /// Raw Boolean field from the stream.
    #[must_use]
    pub fn license_required_raw(&self) -> u32 {
        self.license_required
    }

    /// Interpret the stream's Boolean field as a nonzero/zero value.
    #[must_use]
    pub fn license_required(&self) -> bool {
        self.license_required != 0
    }
}

/// Kind of module item declared by the text-level `PROJECT` stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProjectModuleKind {
    /// `Module=Name`.
    Standard,
    /// `Class=Name`.
    Class,
    /// `Document=Name/&Hxxxxxxxx`.
    Document {
        /// Automation server version written after the module name.
        type_library_version: u32,
    },
    /// `BaseClass=Name`.
    Designer,
}

/// A host-extender reference from the `[Host Extender Info]` section.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProjectHostExtender {
    index: u32,
    guid: String,
    lib_name: String,
    creation_flags: u32,
}

impl ProjectHostExtender {
    /// Construct a host-extender record.
    #[must_use]
    pub fn new(
        index: u32,
        guid: impl Into<String>,
        lib_name: impl Into<String>,
        creation_flags: u32,
    ) -> Self {
        Self {
            index,
            guid: guid.into(),
            lib_name: lib_name.into(),
            creation_flags,
        }
    }

    /// Host-extender index.
    #[must_use]
    pub fn index(&self) -> u32 {
        self.index
    }

    /// Automation type-library GUID.
    #[must_use]
    pub fn guid(&self) -> &str {
        &self.guid
    }

    /// Host-provided Automation type-library name.
    #[must_use]
    pub fn lib_name(&self) -> &str {
        &self.lib_name
    }

    /// Host-provided creation flags.
    #[must_use]
    pub fn creation_flags(&self) -> u32 {
        self.creation_flags
    }
}

/// Coordinates and state of one `PROJECT` editor window.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProjectWindowState {
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
    state: String,
}

impl ProjectWindowState {
    /// Construct a window state. `state` may contain the MS-OVBA `C`, `Z`, and
    /// `I` flags in any order.
    #[must_use]
    pub fn new(left: i32, top: i32, right: i32, bottom: i32, state: impl Into<String>) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
            state: state.into(),
        }
    }

    /// Left coordinate.
    #[must_use]
    pub fn left(&self) -> i32 {
        self.left
    }

    /// Top coordinate.
    #[must_use]
    pub fn top(&self) -> i32 {
        self.top
    }

    /// Right coordinate.
    #[must_use]
    pub fn right(&self) -> i32 {
        self.right
    }

    /// Bottom coordinate.
    #[must_use]
    pub fn bottom(&self) -> i32 {
        self.bottom
    }

    /// Window-state flags (`C`, `Z`, and/or `I`).
    #[must_use]
    pub fn state(&self) -> &str {
        &self.state
    }
}

/// One module's code window and optional designer window from `PROJECT`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProjectWindow {
    module_name: String,
    code: ProjectWindowState,
    designer: Option<ProjectWindowState>,
}

impl ProjectWindow {
    /// Construct a project window record.
    #[must_use]
    pub fn new(
        module_name: impl Into<String>,
        code: ProjectWindowState,
        designer: Option<ProjectWindowState>,
    ) -> Self {
        Self {
            module_name: module_name.into(),
            code,
            designer,
        }
    }

    /// Module associated with the window state.
    #[must_use]
    pub fn module_name(&self) -> &str {
        &self.module_name
    }

    /// Code-editor window state.
    #[must_use]
    pub fn code(&self) -> &ProjectWindowState {
        &self.code
    }

    /// Optional designer-editor window state.
    #[must_use]
    pub fn designer(&self) -> Option<&ProjectWindowState> {
        self.designer.as_ref()
    }
}

/// A typed record from the uncompressed `PROJECT` stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ProjectTextRecord {
    /// `ID="{...}"`.
    Id(String),
    /// A module declaration (`Document=`, `Module=`, `Class=`, or
    /// `BaseClass=`).
    Module {
        /// Module identifier.
        name: String,
        /// Declared module category.
        kind: ProjectModuleKind,
    },
    /// `Package={...}`.
    Package(String),
    /// `HelpFile="..."`.
    HelpFile(String),
    /// `ExeName32="..."`.
    ExeName32(String),
    /// `Name="..."`.
    Name(String),
    /// `HelpContextID="..."`.
    HelpContextId(i32),
    /// `Description="..."`.
    Description(String),
    /// `VersionCompatible32="393222000"`.
    VersionCompatible32(String),
    /// `CMG="..."`.
    ProtectionState(String),
    /// `DPB="..."`.
    Password(String),
    /// `GC="..."`.
    VisibilityState(String),
    /// `[Host Extender Info]`.
    HostExtenderSection,
    /// One host-extender record.
    HostExtender(ProjectHostExtender),
    /// `[Workspace]`.
    WorkspaceSection,
    /// One module editor-window record.
    Window(ProjectWindow),
    /// A blank line retained from the source.
    Blank,
    /// A producer extension or otherwise unrecognized line. The line is
    /// retained for source-preserving edits.
    Unknown(String),
}

/// Bounded, typed view of the `PROJECT` text stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProjectText {
    raw: String,
    records: Vec<ProjectTextRecord>,
    edited: bool,
}

/// One finite property assignment from a `VBFrame` designer stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DesignerProperty {
    name: String,
    value: String,
    comment: Option<String>,
}

impl DesignerProperty {
    /// Construct a designer property with its textual value and optional
    /// source comment.
    #[must_use]
    pub fn new(name: impl Into<String>, value: impl Into<String>, comment: Option<String>) -> Self {
        Self {
            name: name.into(),
            value: value.into(),
            comment,
        }
    }

    /// Property name from the finite VBFrame vocabulary.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Original textual property value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Optional source comment following the property value.
    #[must_use]
    pub fn comment(&self) -> Option<&str> {
        self.comment.as_deref()
    }
}

/// Parsed inert metadata from one designer's `VBFrame` stream.
#[derive(Debug, PartialEq, Eq)]
pub struct DesignerFrame {
    module_name: String,
    class_id: String,
    properties: Vec<DesignerProperty>,
    source: Text,
}

impl DesignerFrame {
    /// Name of the associated designer module.
    #[must_use]
    pub fn module_name(&self) -> &str {
        &self.module_name
    }

    /// Designer CLSID in canonical textual GUID form.
    #[must_use]
    pub fn class_id(&self) -> &str {
        &self.class_id
    }

    /// Typed properties in source order.
    #[must_use]
    pub fn properties(&self) -> &[DesignerProperty] {
        &self.properties
    }

    /// Decoded source text, retained for exact passive round-tripping.
    #[must_use]
    pub fn source(&self) -> &Text {
        &self.source
    }

    /// Construct and validate a frame from raw MBCS bytes.
    pub fn from_raw(
        module_name: impl Into<String>,
        raw: Vec<u8>,
        page: Mbcs,
        limits: &Limits,
    ) -> Result<Self, Error> {
        check_limit(
            "VBA designer stream bytes",
            raw.len(),
            limits.max_decompressed_stream_bytes,
        )?;
        let module_name = module_name.into();
        validate_project_module_identifier(&module_name)?;
        let source = Text::decode(raw, page);
        parse_vbframe(module_name, source, limits)
    }
}

impl ProjectText {
    /// Parse a decoded `PROJECT` stream while retaining its exact text.
    pub fn parse(text: &str, limits: &Limits) -> Result<Self, Error> {
        check_limit(
            "PROJECT text bytes",
            text.len(),
            limits.max_decompressed_stream_bytes,
        )?;
        let mut records = Vec::new();
        let mut section = ProjectTextSection::Properties;
        for line in project_lines(text)? {
            records
                .try_reserve(1)
                .map_err(|_| invalid("PROJECT text record allocation failed"))?;
            check_limit(
                "PROJECT text line bytes",
                line.len(),
                limits.max_string_bytes,
            )?;
            let trimmed = line.trim_matches([' ', '\t']);
            if trimmed.is_empty() {
                records.push(ProjectTextRecord::Blank);
                continue;
            }
            if trimmed.eq_ignore_ascii_case("[Host Extender Info]") {
                section = ProjectTextSection::HostExtenders;
                records.push(ProjectTextRecord::HostExtenderSection);
                continue;
            }
            if trimmed.eq_ignore_ascii_case("[Workspace]") {
                section = ProjectTextSection::Workspace;
                records.push(ProjectTextRecord::WorkspaceSection);
                continue;
            }
            let record = match section {
                ProjectTextSection::HostExtenders => parse_host_extender_line(trimmed),
                ProjectTextSection::Workspace => parse_workspace_line(trimmed),
                ProjectTextSection::Properties => parse_project_property_line(trimmed, limits),
            }?;
            if matches!(&record, ProjectTextRecord::Unknown(_)) && line != trimmed {
                records.push(ProjectTextRecord::Unknown(line.to_owned()));
            } else {
                records.push(record);
            }
        }
        Ok(Self {
            raw: text.to_owned(),
            records,
            edited: false,
        })
    }

    /// Exact decoded source text before any edit.
    #[must_use]
    pub fn raw(&self) -> &str {
        &self.raw
    }

    /// Ordered typed records.
    #[must_use]
    pub fn records(&self) -> &[ProjectTextRecord] {
        &self.records
    }

    pub(crate) fn validate_payload_structure(&self, directory: &Dir) -> Result<(), Error> {
        validate_vba_identifier(directory.project_name(), "PROJECTNAME")?;
        let records = &self.records;
        let mut index = 0usize;
        let next = |index: &mut usize| {
            let record = records.get(*index);
            *index = index.saturating_add(1);
            record
        };

        if !matches!(next(&mut index), Some(ProjectTextRecord::Id(_))) {
            return Err(invalid(
                "PROJECT stream must begin with exactly one ID record",
            ));
        }
        while let Some(record) = records.get(index) {
            match record {
                ProjectTextRecord::Module { .. } | ProjectTextRecord::Package(_) => {
                    index += 1;
                },
                _ => break,
            }
        }
        if let Some(ProjectTextRecord::HelpFile(_)) = records.get(index) {
            index += 1;
        }
        if let Some(ProjectTextRecord::ExeName32(_)) = records.get(index) {
            index += 1;
        }
        let Some(ProjectTextRecord::Name(name)) = next(&mut index) else {
            return Err(invalid("PROJECT stream is missing its Name record"));
        };
        if name != directory.project_name() {
            return Err(invalid(
                "PROJECT Name does not match the dir PROJECTNAME record",
            ));
        }
        if !matches!(next(&mut index), Some(ProjectTextRecord::HelpContextId(_))) {
            return Err(invalid(
                "PROJECT stream is missing its HelpContextID record",
            ));
        }
        if matches!(records.get(index), Some(ProjectTextRecord::Description(_))) {
            index += 1;
        }
        if matches!(
            records.get(index),
            Some(ProjectTextRecord::VersionCompatible32(_))
        ) {
            index += 1;
        }
        if !matches!(
            next(&mut index),
            Some(ProjectTextRecord::ProtectionState(_))
        ) {
            return Err(invalid(
                "PROJECT stream is missing its CMG protection record",
            ));
        }
        if !matches!(next(&mut index), Some(ProjectTextRecord::Password(_))) {
            return Err(invalid("PROJECT stream is missing its DPB password record"));
        }
        if !matches!(
            next(&mut index),
            Some(ProjectTextRecord::VisibilityState(_))
        ) {
            return Err(invalid(
                "PROJECT stream is missing its GC visibility record",
            ));
        }
        if !matches!(next(&mut index), Some(ProjectTextRecord::Blank)) {
            return Err(invalid(
                "PROJECT stream is missing the separator before Host Extender Info",
            ));
        }
        if !matches!(
            next(&mut index),
            Some(ProjectTextRecord::HostExtenderSection)
        ) {
            return Err(invalid(
                "PROJECT stream is missing its Host Extender Info section",
            ));
        }

        let mut host_indices = Vec::new();
        while let Some(record) = records.get(index) {
            match record {
                ProjectTextRecord::HostExtender(extender) => {
                    if host_indices.contains(&extender.index()) {
                        return Err(invalid("PROJECT host extender index is duplicated"));
                    }
                    host_indices.push(extender.index());
                    index += 1;
                },
                ProjectTextRecord::Blank => {
                    index += 1;
                    if !matches!(
                        records.get(index),
                        Some(ProjectTextRecord::WorkspaceSection)
                    ) {
                        return Err(invalid(
                            "PROJECT stream has a blank record outside its workspace separator",
                        ));
                    }
                    break;
                },
                ProjectTextRecord::WorkspaceSection => {
                    return Err(invalid(
                        "PROJECT Workspace is missing the separator after Host Extender Info",
                    ));
                },
                _ => {
                    return Err(invalid(
                        "PROJECT stream contains a record outside Host Extender Info",
                    ));
                },
            }
        }
        if matches!(
            records.get(index),
            Some(ProjectTextRecord::WorkspaceSection)
        ) {
            index += 1;
            let mut windows = Vec::new();
            while let Some(record) = records.get(index) {
                let ProjectTextRecord::Window(window) = record else {
                    return Err(invalid(
                        "PROJECT stream contains a non-window record in Workspace",
                    ));
                };
                if windows
                    .iter()
                    .any(|name: &&str| name.eq_ignore_ascii_case(window.module_name()))
                {
                    return Err(invalid(
                        "PROJECT workspace contains a duplicate module window",
                    ));
                }
                let project_module_count = records
                    .iter()
                    .filter(|record| {
                        matches!(
                            record,
                            ProjectTextRecord::Module { name, .. }
                                if name.eq_ignore_ascii_case(window.module_name())
                        )
                    })
                    .count();
                if project_module_count != 1 {
                    return Err(invalid(format!(
                        "PROJECT workspace window {} does not have exactly one PROJECT module declaration",
                        window.module_name()
                    )));
                }
                let directory_module_count = directory
                    .modules()
                    .iter()
                    .filter(|module| module.name().eq_ignore_ascii_case(window.module_name()))
                    .count();
                if directory_module_count != 1 {
                    return Err(invalid(format!(
                        "PROJECT workspace window {} does not have exactly one dir MODULE record",
                        window.module_name()
                    )));
                }
                windows.push(window.module_name());
                index += 1;
            }
        }
        if index != records.len() {
            return Err(invalid(
                "PROJECT stream contains trailing or unknown records",
            ));
        }

        for (name, kind) in records.iter().filter_map(|record| match record {
            ProjectTextRecord::Module { name, kind } => Some((name, kind)),
            _ => None,
        }) {
            let matching_count = directory
                .modules()
                .iter()
                .filter(|module| module.name().eq_ignore_ascii_case(name))
                .count();
            if matching_count != 1 {
                return Err(invalid(format!(
                    "PROJECT module {name} does not have exactly one dir MODULE record"
                )));
            }
            let matching = directory
                .modules()
                .iter()
                .find(|module| module.name().eq_ignore_ascii_case(name))
                .ok_or_else(|| {
                    invalid(format!("PROJECT module {name} has no dir MODULE record"))
                })?;
            let compatible = match kind {
                ProjectModuleKind::Standard => matching.kind() == Kind::Procedural,
                ProjectModuleKind::Class
                | ProjectModuleKind::Document { .. }
                | ProjectModuleKind::Designer => matching.kind() == Kind::DocumentClassOrDesigner,
            };
            if !compatible {
                return Err(invalid(format!(
                    "PROJECT module {name} has an incompatible dir MODULETYPE"
                )));
            }
        }
        for module in directory.modules() {
            let matching = records
                .iter()
                .filter(|record| {
                    matches!(
                        record,
                        ProjectTextRecord::Module { name, .. }
                            if name.eq_ignore_ascii_case(module.name())
                    )
                })
                .count();
            if matching != 1 {
                return Err(invalid(format!(
                    "dir MODULE {} does not have exactly one PROJECT module declaration",
                    module.name()
                )));
            }
        }
        Ok(())
    }

    /// Project name, if a `Name` record was present.
    #[must_use]
    pub fn project_name(&self) -> Option<&str> {
        self.records.iter().find_map(|record| match record {
            ProjectTextRecord::Name(name) => Some(name.as_str()),
            _ => None,
        })
    }

    /// Replace the first project `Name` record and mark the view edited.
    pub fn set_project_name(&mut self, name: impl Into<String>) -> Result<(), Error> {
        let name = name.into();
        if name.is_empty()
            || name.chars().count() > 128
            || name
                .chars()
                .any(|character| !valid_quoted_character(character))
        {
            return Err(invalid(
                "PROJECT Name is outside its 1..=128 character bound",
            ));
        }
        let Some(record) = self
            .records
            .iter_mut()
            .find(|record| matches!(record, ProjectTextRecord::Name(_)))
        else {
            return Err(invalid("PROJECT stream has no Name record"));
        };
        if matches!(record, ProjectTextRecord::Name(current) if current == &name) {
            return Ok(());
        }
        *record = ProjectTextRecord::Name(name);
        self.edited = true;
        Ok(())
    }

    /// Replace a module identifier throughout module declarations and window
    /// records, preserving their order and all unknown source lines.
    pub fn rename_module(&mut self, old: &str, new: impl Into<String>) -> Result<(), Error> {
        let new = new.into();
        validate_project_module_identifier(&new)?;
        let mut found = false;
        let changed = old != new;
        for record in &mut self.records {
            match record {
                ProjectTextRecord::Module { name, .. } if name == old => {
                    if changed {
                        *name = new.clone();
                    }
                    found = true;
                },
                ProjectTextRecord::Window(window) if window.module_name == old => {
                    if changed {
                        window.module_name = new.clone();
                    }
                    found = true;
                },
                _ => {},
            }
        }
        if !found {
            return Err(invalid(format!("PROJECT stream has no module named {old}")));
        }
        self.edited |= changed;
        Ok(())
    }

    /// Render the source. Unedited views return the exact original text;
    /// edited views use canonical CRLF records while retaining unknown lines.
    #[must_use]
    pub fn to_text(&self) -> String {
        if !self.edited {
            return self.raw.clone();
        }
        let mut output = String::new();
        for record in &self.records {
            render_project_text_record(record, &mut output);
        }
        output
            .replace("\r\n", "\n")
            .replace("\n\r", "\n")
            .replace('\n', "\r\n")
    }

    /// Encode the rendered text using the project's checked code page.
    pub fn encode(&self, page: Mbcs, limits: &Limits) -> Result<Vec<u8>, Error> {
        let text = self.to_text();
        // Public record variants can be assembled by callers, so validate the
        // rendered grammar before encoding it. A source-bound unedited view is
        // still returned byte-for-byte by Project::encode_project_text.
        let _ = Self::parse(&text, limits)?;
        check_limit(
            "PROJECT text bytes",
            text.len(),
            limits.max_decompressed_stream_bytes,
        )?;
        let encoded = super::dir::encode_mbcs(&text, page, "PROJECT stream")?;
        check_limit(
            "PROJECT stream bytes",
            encoded.len(),
            limits.max_decompressed_stream_bytes,
        )?;
        Ok(encoded)
    }
}

/// Raw and decoded text from an MS-OVBA stream.
#[derive(Debug, PartialEq, Eq)]
pub struct Text {
    raw: Vec<u8>,
    decoded: String,
    had_decode_errors: bool,
}

impl Text {
    /// Original bytes after decompression, before character decoding.
    #[must_use]
    pub fn raw(&self) -> &[u8] {
        &self.raw
    }

    /// Text decoded with the project's declared code page.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.decoded
    }

    /// Whether malformed byte sequences were replaced while decoding.
    #[must_use]
    pub fn had_decode_errors(&self) -> bool {
        self.had_decode_errors
    }

    fn decode(raw: Vec<u8>, page: Mbcs) -> Self {
        let (recovered, had_decode_errors) = page.recover(&raw);
        let decoded = recovered.into_owned();
        Self {
            raw,
            decoded,
            had_decode_errors,
        }
    }
}

/// One inert VBA module and its typed directory metadata.
#[derive(Debug, PartialEq, Eq)]
pub struct Module {
    name: String,
    stream_name: String,
    text_offset: u32,
    kind: Kind,
    read_only: bool,
    private: bool,
    source: Text,
}

impl Module {
    /// VBA identifier for this module.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// CFB stream containing this module.
    #[must_use]
    pub fn stream_name(&self) -> &str {
        &self.stream_name
    }

    /// Byte offset at which compressed source begins in the module stream.
    #[must_use]
    pub fn text_offset(&self) -> u32 {
        self.text_offset
    }

    /// Broad module category from the `dir` stream.
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Whether this module is marked read-only.
    #[must_use]
    pub fn is_read_only(&self) -> bool {
        self.read_only
    }

    /// Whether this module is private to its project.
    #[must_use]
    pub fn is_private(&self) -> bool {
        self.private
    }

    /// Decompressed, inert module source.
    #[must_use]
    pub fn source(&self) -> &Text {
        &self.source
    }
}

/// A complete inert VBA project loaded from a CFB project-root storage.
#[derive(Debug, PartialEq, Eq)]
pub struct Project {
    root_path: Vec<String>,
    page: Mbcs,
    name: String,
    references: Vec<Reference>,
    dir_source: Vec<u8>,
    dir_compressed_source: Vec<u8>,
    properties: Text,
    project_text_cache: OnceLock<ProjectText>,
    project_wm_source: Option<Vec<u8>>,
    project_wm_cache: OnceLock<Vec<NameMap>>,
    license_info: Option<Vec<LicenseInfo>>,
    project_lk_source: Option<Vec<u8>>,
    designer_frames: Vec<DesignerFrame>,
    modules: Vec<Module>,
}

impl Project {
    /// Parse a standalone `vbaProject.bin` payload with safe default limits.
    ///
    /// The CFB bytes are borrowed and never copied.
    ///
    /// # Errors
    ///
    /// Returns an error if the payload is not a readable CFB file, the project
    /// structures are malformed, or the default resource limits are exceeded.
    pub fn read(bytes: &[u8]) -> Result<Self, Error> {
        Self::read_with(bytes, &Limits::default())
    }

    /// Parse a standalone `vbaProject.bin` payload with explicit limits.
    ///
    /// This is the container-independent entry point for OOXML hosts. The
    /// borrowed payload is retained only for the duration of parsing; the
    /// returned project owns its bounded semantic metadata and inert source.
    ///
    /// # Errors
    ///
    /// Returns an error if the payload is not a readable CFB file, the project
    /// structures are malformed, or `limits` are exceeded.
    pub fn read_with(bytes: &[u8], limits: &Limits) -> Result<Self, Error> {
        check_limit(
            "standalone VBA CFB bytes",
            bytes.len(),
            limits.max_cfb_bytes,
        )?;
        let mut ole = OleFile::open(Cursor::new(bytes))?;
        Self::open(&mut ole, &[], limits)
    }

    /// Load an MS-OVBA project rooted at `project_root_path`.
    ///
    /// For an OOXML `vbaProject.bin` part, the project root is normally the
    /// CFB root and this path is empty. Legacy Excel normally passes
    /// `["_VBA_PROJECT_CUR"]`; other hosts use the storage discovered from
    /// their CFB directory.
    ///
    /// # Errors
    ///
    /// Returns an error if required streams are missing from `ole`, the
    /// project structures are malformed, or `limits` are exceeded.
    pub fn open<R: Read + Seek>(
        ole: &mut OleFile<R>,
        project_root_path: &[&str],
        limits: &Limits,
    ) -> Result<Self, Error> {
        let mut version_path = project_root_path.to_vec();
        version_path.extend([VBA_STORAGE_NAME, VERSION_PROJECT_STREAM_NAME]);
        let version_stream =
            read_limited_stream(ole, &version_path, limits.max_compressed_stream_bytes)?;
        if version_stream.len() < VERSION_PROJECT_HEADER_BYTES {
            return Err(invalid(
                "_VBA_PROJECT stream is shorter than its seven-byte header",
            ));
        }

        let mut dir_path = project_root_path.to_vec();
        dir_path.extend([VBA_STORAGE_NAME, DIR_STREAM_NAME]);
        let compressed_dir =
            read_limited_stream(ole, &dir_path, limits.max_compressed_stream_bytes)?;
        let dir_source = codec::decode(&compressed_dir, limits)?;
        let directory = Dir::parse_decompressed(&dir_source, limits)?;
        let page = directory.page();

        let mut project_wm_path = project_root_path.to_vec();
        project_wm_path.push(PROJECT_WM_STREAM_NAME);
        let project_wm_source = read_optional_limited_stream(
            ole,
            &project_wm_path,
            limits.max_decompressed_stream_bytes,
        )?;

        let mut project_lk_path = project_root_path.to_vec();
        project_lk_path.push(PROJECT_LK_STREAM_NAME);
        let project_lk_source = read_optional_limited_stream(
            ole,
            &project_lk_path,
            limits.max_decompressed_stream_bytes,
        )?;
        let license_info = project_lk_source
            .as_deref()
            .map(|source| parse_project_lk(source, limits))
            .transpose()?;

        let mut project_path = project_root_path.to_vec();
        project_path.push(PROJECT_STREAM_NAME);
        let project_properties = Text::decode(
            read_limited_stream(ole, &project_path, limits.max_decompressed_stream_bytes)?,
            page,
        );

        let mut modules = Vec::with_capacity(directory.modules().len());
        let mut total_source_bytes = 0usize;
        for metadata in directory.modules() {
            let mut module_path = project_root_path.to_vec();
            module_path.extend([VBA_STORAGE_NAME, metadata.stream_name()]);
            let stream =
                read_limited_stream(ole, &module_path, limits.max_compressed_stream_bytes)?;
            let Ok(text_offset) = usize::try_from(metadata.text_offset()) else {
                return Err(invalid("module text offset does not fit usize"));
            };
            let compressed_source = stream.get(text_offset..).ok_or_else(|| {
                invalid(format!(
                    "module {} text offset {} exceeds stream size {}",
                    metadata.name(),
                    text_offset,
                    stream.len()
                ))
            })?;
            // Bound each decompression by the aggregate source budget still
            // available. Without this preflight, a later module could
            // allocate a full per-stream maximum before the aggregate check
            // rejected it.
            let remaining_source_bytes = limits
                .max_total_source_bytes
                .saturating_sub(total_source_bytes);
            let stream_maximum = limits
                .max_decompressed_stream_bytes
                .min(remaining_source_bytes);
            let stream_limits = Limits {
                max_decompressed_stream_bytes: stream_maximum,
                ..*limits
            };
            let source_bytes = match codec::decode(compressed_source, &stream_limits) {
                Ok(bytes) => bytes,
                Err(Error::LimitExceeded {
                    limit: "decompressed VBA stream bytes",
                    actual,
                    ..
                }) if stream_maximum < limits.max_decompressed_stream_bytes => {
                    return Err(Error::LimitExceeded {
                        limit: "aggregate VBA module source bytes",
                        actual: total_source_bytes.saturating_add(actual),
                        maximum: limits.max_total_source_bytes,
                    });
                },
                Err(error) => return Err(error),
            };
            total_source_bytes = total_source_bytes
                .checked_add(source_bytes.len())
                .ok_or_else(|| invalid("aggregate VBA source size overflow"))?;
            check_limit(
                "aggregate VBA module source bytes",
                total_source_bytes,
                limits.max_total_source_bytes,
            )?;
            modules.push(Module {
                name: metadata.name().to_owned(),
                stream_name: metadata.stream_name().to_owned(),
                text_offset: metadata.text_offset(),
                kind: metadata.kind(),
                read_only: metadata.is_read_only(),
                private: metadata.is_private(),
                source: Text::decode(source_bytes, page),
            });
        }
        let designer_frames = read_designer_frames(ole, project_root_path, &modules, page, limits)?;

        Ok(Self {
            root_path: project_root_path
                .iter()
                .map(|component| (*component).to_owned())
                .collect(),
            page,
            name: directory.project_name().to_owned(),
            references: directory.references().to_vec(),
            dir_source,
            dir_compressed_source: compressed_dir,
            properties: project_properties,
            project_text_cache: OnceLock::new(),
            project_wm_source,
            project_wm_cache: OnceLock::new(),
            license_info,
            project_lk_source,
            designer_frames,
            modules,
        })
    }

    /// CFB path of the MS-OVBA project root.
    #[must_use]
    pub fn project_root_path(&self) -> &[String] {
        &self.root_path
    }

    /// Checked project page used to decode MBCS text.
    #[must_use]
    pub fn page(&self) -> Mbcs {
        self.page
    }

    /// Numeric project-page identifier.
    #[must_use]
    pub fn page_id(&self) -> u16 {
        self.page.id16()
    }

    /// VBA project identifier.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// External references declared by the project, in `dir` order.
    #[must_use]
    pub fn references(&self) -> &[Reference] {
        &self.references
    }

    /// Exact compressed `dir` stream bytes from the source project.
    #[must_use]
    pub fn dir_raw(&self) -> &[u8] {
        &self.dir_compressed_source
    }

    /// Re-encode the source-bound `PROJECTREFERENCES` array in the `dir`
    /// stream.  Unchanged references return the exact source-compressed bytes;
    /// a changed reference rewrites only that array and retains all project
    /// information, module records, reserved values, and unknown bytes.
    pub fn encode_dir_with_references(
        &self,
        references: &[Reference],
        limits: &Limits,
    ) -> Result<Vec<u8>, Error> {
        if references == self.references {
            // Reparse the source-bound directory under the caller's current
            // limits before returning the exact compressed bytes.  The
            // no-op path must not bypass string, reference, or module bounds.
            let _ = Dir::parse_decompressed(&self.dir_source, limits)?;
            check_limit(
                "compressed VBA stream bytes",
                self.dir_compressed_source.len(),
                limits.max_compressed_stream_bytes,
            )?;
            check_limit(
                "decompressed VBA stream bytes",
                self.dir_source.len(),
                limits.max_decompressed_stream_bytes,
            )?;
            return Ok(self.dir_compressed_source.clone());
        }
        encode_dir_with_references(&self.dir_source, references, limits)
    }

    /// Decoded text of the uncompressed `PROJECT` stream.
    #[must_use]
    pub fn project_properties(&self) -> &Text {
        &self.properties
    }

    /// Typed records from the uncompressed `PROJECT` stream.
    ///
    /// Parsing is deferred until this accessor is called so opening a project
    /// remains compatible with producer extensions represented as unknown
    /// text lines. Successful parses are cached; malformed parses are not.
    pub fn project_text(&self, limits: &Limits) -> Result<&ProjectText, Error> {
        check_limit("VBA module count", self.modules.len(), limits.max_modules)?;
        if let Some(cache) = self.project_text_cache.get() {
            // A semantic cache is reusable only after it has been checked
            // against the caller's current limits. In particular, a caller
            // may deliberately tighten max_string_bytes after a previous
            // unrestricted accessor populated the cache.
            let _ = ProjectText::parse(self.properties.text(), limits)?;
            return Ok(cache);
        }
        let parsed = ProjectText::parse(self.properties.text(), limits)?;
        let _ = self.project_text_cache.set(parsed);
        self.project_text_cache
            .get()
            .ok_or_else(|| invalid("PROJECT text cache was not initialized"))
    }

    /// Encode a `PROJECT` text view for a source-bound stream replacement.
    ///
    /// An unedited view obtained from this project returns the exact source
    /// bytes, including any producer code-page representation.  Edited views
    /// are encoded using the checked project code page and retain their
    /// ordered unknown records through [`ProjectText::to_text`].
    pub fn encode_project_text(
        &self,
        text: &ProjectText,
        limits: &Limits,
    ) -> Result<Vec<u8>, Error> {
        check_limit("VBA module count", self.modules.len(), limits.max_modules)?;
        if !text.edited && text.raw == self.properties.text() {
            // The exact-source fast path still validates the complete typed
            // grammar and per-record limits for this caller. Returning the
            // original bytes must not bypass a newly tighter limit.
            let _ = ProjectText::parse(self.properties.text(), limits)?;
            check_limit(
                "PROJECT text bytes",
                self.properties.raw.len(),
                limits.max_decompressed_stream_bytes,
            )?;
            return Ok(self.properties.raw.clone());
        }
        text.encode(self.page, limits)
    }

    /// Optional MBCS/UTF-16 module-name map from the `PROJECTwm` stream.
    #[must_use = "PROJECTwm parsing result should be checked"]
    pub fn project_wm(&self, limits: &Limits) -> Result<Option<&[NameMap]>, Error> {
        check_limit("VBA module count", self.modules.len(), limits.max_modules)?;
        let Some(source) = self.project_wm_source.as_deref() else {
            return Ok(None);
        };
        if let Some(cache) = self.project_wm_cache.get() {
            // Revalidate the source under the caller's limits even when the
            // parsed map is already cached. This keeps cached and cold paths
            // equivalent for bounded callers.
            let _ = parse_project_wm(source, self.page, &self.modules, limits)?;
            return Ok(Some(cache));
        }
        let parsed = parse_project_wm(source, self.page, &self.modules, limits)?;
        let _ = self.project_wm_cache.set(parsed);
        Ok(self.project_wm_cache.get().map(Vec::as_slice))
    }

    /// Exact bytes of the optional `PROJECTwm` stream, when present.
    #[must_use]
    pub fn project_wm_raw(&self) -> Option<&[u8]> {
        self.project_wm_source.as_deref()
    }

    /// Encode an ordered `PROJECTwm` edit against this project's directory.
    ///
    /// The map count, order, and Unicode spelling are checked against the
    /// `MODULE` records from `dir`; the MBCS spelling may change only when it
    /// still decodes to the same Unicode module identifier.  Passing the
    /// parsed maps back unchanged returns the original stream bytes, which
    /// preserves a producer's exact source representation for a no-op edit.
    pub fn encode_project_wm(&self, maps: &[NameMap], limits: &Limits) -> Result<Vec<u8>, Error> {
        check_limit("VBA module count", self.modules.len(), limits.max_modules)?;
        if let Some(source) = self.project_wm_source.as_deref() {
            let parsed = self
                .project_wm(limits)?
                .ok_or_else(|| invalid("PROJECTwm source disappeared while encoding an edit"))?;
            if parsed == maps {
                return Ok(source.to_vec());
            }
        }
        encode_project_wm(maps, self.page, &self.modules, limits)
    }

    /// ActiveX control license metadata from the optional `PROJECTlk` stream.
    #[must_use]
    pub fn license_info(&self) -> Option<&[LicenseInfo]> {
        self.license_info.as_deref()
    }

    /// Exact bytes of the optional `PROJECTlk` stream, when present.
    #[must_use]
    pub fn project_lk_raw(&self) -> Option<&[u8]> {
        self.project_lk_source.as_deref()
    }

    /// Parsed finite metadata from designer `VBFrame` streams.
    #[must_use]
    pub fn designer_frames(&self) -> &[DesignerFrame] {
        &self.designer_frames
    }

    /// Validate the cross-stream ownership relation for designer metadata.
    ///
    /// A VBFrame stream is owned by exactly one BaseClass declaration in
    /// PROJECT, and every BaseClass declaration must have exactly one frame.
    /// The low-level Project::open reader remains source-compatible and can
    /// expose a partially observed project; owned Payload admission calls this
    /// check before retaining the input bytes.
    pub(crate) fn validate_designer_ownership(
        &self,
        project_text: &ProjectText,
    ) -> Result<(), Error> {
        for frame in &self.designer_frames {
            let declarations = project_text
                .records()
                .iter()
                .filter_map(|record| match record {
                    ProjectTextRecord::Module {
                        name,
                        kind: ProjectModuleKind::Designer,
                    } if name.eq_ignore_ascii_case(frame.module_name()) => Some(name),
                    _ => None,
                })
                .count();
            if declarations != 1 {
                return Err(invalid(format!(
                    "VBFrame for {} has no unique BaseClass declaration",
                    frame.module_name()
                )));
            }
        }
        for record in project_text.records() {
            let ProjectTextRecord::Module {
                name,
                kind: ProjectModuleKind::Designer,
            } = record
            else {
                continue;
            };
            let matching_frames = self
                .designer_frames
                .iter()
                .filter(|frame| frame.module_name().eq_ignore_ascii_case(name))
                .count();
            if matching_frames != 1 {
                return Err(invalid(format!(
                    "BaseClass {name} must have exactly one VBFrame stream"
                )));
            }
            let Some(module) = self
                .modules
                .iter()
                .find(|module| module.name().eq_ignore_ascii_case(name))
            else {
                return Err(invalid(format!(
                    "BaseClass {name} has no matching directory module"
                )));
            };
            if module.kind() != Kind::DocumentClassOrDesigner {
                return Err(invalid(format!(
                    "BaseClass {name} has an incompatible directory module kind"
                )));
            }
        }
        Ok(())
    }

    pub(crate) fn validate_payload(&self, limits: &Limits) -> Result<(), Error> {
        let project_text = self.project_text(limits)?;
        let directory = Dir::parse_decompressed_strict(&self.dir_source, limits)?;
        project_text.validate_payload_structure(&directory)?;
        self.validate_designer_ownership(project_text)
    }

    /// Modules in `dir`-stream order.
    #[must_use]
    pub fn modules(&self) -> &[Module] {
        &self.modules
    }
}

fn read_limited_stream<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: &[&str],
    maximum: usize,
) -> Result<Vec<u8>, Error> {
    let (parent, stream_name) = path
        .split_last()
        .map(|(last, parent)| (parent, *last))
        .ok_or_else(|| invalid("VBA stream path must not be empty"))?;
    let entry = ole
        .list_directory_entries(parent)?
        .into_iter()
        .find(|entry| entry.name.eq_ignore_ascii_case(stream_name))
        .ok_or(OleError::StreamNotFound)?;
    let Ok(size) = usize::try_from(entry.size) else {
        return Err(invalid("VBA stream size does not fit usize"));
    };
    check_limit("VBA CFB stream bytes", size, maximum)?;
    let stream = ole.open_stream(path)?;
    check_limit("VBA CFB stream bytes", stream.len(), maximum)?;
    Ok(stream)
}

fn read_optional_limited_stream<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: &[&str],
    maximum: usize,
) -> Result<Option<Vec<u8>>, Error> {
    let (parent, stream_name) = path
        .split_last()
        .map(|(last, parent)| (parent, *last))
        .ok_or_else(|| invalid("VBA stream path must not be empty"))?;
    let exists = ole
        .list_directory_entries(parent)?
        .into_iter()
        .any(|entry| entry.name.eq_ignore_ascii_case(stream_name));
    if !exists {
        return Ok(None);
    }
    read_limited_stream(ole, path, maximum).map(Some)
}

fn parse_project_wm(
    bytes: &[u8],
    page: Mbcs,
    modules: &[Module],
    limits: &Limits,
) -> Result<Vec<NameMap>, Error> {
    check_limit("VBA module count", modules.len(), limits.max_modules)?;
    check_limit(
        "decompressed VBA stream bytes",
        bytes.len(),
        limits.max_decompressed_stream_bytes,
    )?;
    // Each module contributes at least one non-empty MBCS byte plus its
    // NUL, one UTF-16 code unit plus its NUL, and the stream has a final
    // UTF-16 NUL terminator.  Reject a structurally impossible stream before
    // reserving the bounded map vector.  This keeps malformed count fields
    // from turning a one-byte input into a count-sized allocation.
    let minimum_bytes = modules
        .len()
        .checked_mul(6)
        .and_then(|length| length.checked_add(2))
        .ok_or_else(|| invalid("PROJECTwm minimum length overflow"))?;
    if bytes.len() < minimum_bytes {
        return Err(invalid(format!(
            "PROJECTwm stream is too short for {} module-name pairs",
            modules.len()
        )));
    }
    let mut cursor = 0usize;
    let mut maps = Vec::new();
    maps.try_reserve(modules.len())
        .map_err(|_| invalid("PROJECTwm map allocation failed"))?;
    for module in modules {
        let mbcs_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != 0 {
            cursor += 1;
            check_limit(
                "PROJECTwm module-name bytes",
                cursor.saturating_sub(mbcs_start),
                limits.max_string_bytes,
            )?;
        }
        if cursor == mbcs_start {
            return Err(invalid("PROJECTwm MBCS module name is empty"));
        }
        if cursor >= bytes.len() {
            return Err(invalid("PROJECTwm MBCS module name is unterminated"));
        }
        let mbcs_name = page
            .decode(&bytes[mbcs_start..cursor])
            .map_err(|_| invalid("PROJECTwm MBCS module name is not decodable"))?
            .into_owned();
        cursor += 1;

        let unicode_start = cursor;
        loop {
            let end = cursor
                .checked_add(2)
                .ok_or_else(|| invalid("PROJECTwm Unicode name offset overflow"))?;
            let raw = bytes
                .get(cursor..end)
                .ok_or_else(|| invalid("PROJECTwm Unicode module name is truncated"))?;
            let code_unit = u16::from_le_bytes([raw[0], raw[1]]);
            cursor = end;
            if code_unit == 0 {
                break;
            }
            check_limit(
                "PROJECTwm Unicode module-name bytes",
                cursor.saturating_sub(unicode_start),
                limits.max_string_bytes,
            )?;
        }
        let unicode_bytes = &bytes[unicode_start..cursor - 2];
        let unicode_name = decode_utf16(unicode_bytes, "PROJECTwm Unicode module name")?;
        if unicode_name != mbcs_name || unicode_name != module.name() {
            return Err(invalid(format!(
                "PROJECTwm module name does not match dir module {}",
                module.name()
            )));
        }
        maps.push(NameMap::new(mbcs_name, unicode_name));
    }
    let terminator = bytes
        .get(cursor..cursor.saturating_add(2))
        .ok_or_else(|| invalid("PROJECTwm stream is missing its terminator"))?;
    if terminator != [0, 0] {
        return Err(invalid("PROJECTwm stream terminator must be zero"));
    }
    cursor += 2;
    if cursor != bytes.len() {
        return Err(invalid("PROJECTwm stream has trailing bytes"));
    }
    Ok(maps)
}

fn encode_project_wm(
    maps: &[NameMap],
    page: Mbcs,
    modules: &[Module],
    limits: &Limits,
) -> Result<Vec<u8>, Error> {
    if maps.len() != modules.len() {
        return Err(invalid(format!(
            "PROJECTwm map count {} does not match dir module count {}",
            maps.len(),
            modules.len()
        )));
    }
    let mut output = Vec::new();
    for (index, (map, module)) in maps.iter().zip(modules).enumerate() {
        if map.mbcs_name().is_empty() || map.unicode_name().is_empty() {
            return Err(invalid(format!(
                "PROJECTwm map {index} contains an empty module name"
            )));
        }
        if map.unicode_name() != module.name() {
            return Err(invalid(format!(
                "PROJECTwm map {index} Unicode name does not match dir module {}",
                module.name()
            )));
        }
        check_limit(
            "VBA input string bytes",
            map.mbcs_name().len(),
            limits.max_string_bytes.saturating_mul(4),
        )?;
        let mbcs = super::dir::encode_mbcs(map.mbcs_name(), page, "PROJECTwm module name")?;
        check_limit("VBA string bytes", mbcs.len(), limits.max_string_bytes)?;
        let decoded = page
            .decode(&mbcs)
            .map_err(|_| invalid(format!("PROJECTwm map {index} MBCS name is not decodable")))?;
        if decoded != map.unicode_name() {
            return Err(invalid(format!(
                "PROJECTwm map {index} MBCS and Unicode names differ"
            )));
        }
        output.extend_from_slice(&mbcs);
        output.push(0);

        let unicode: Vec<u16> = map.unicode_name().encode_utf16().collect();
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

fn read_designer_frames<R: Read + Seek>(
    ole: &mut OleFile<R>,
    project_root_path: &[&str],
    modules: &[Module],
    page: Mbcs,
    limits: &Limits,
) -> Result<Vec<DesignerFrame>, Error> {
    let root_len = project_root_path.len();
    let mut frames = Vec::new();
    for path in ole.list_streams() {
        if path.len() != root_len.saturating_add(2)
            || path
                .iter()
                .take(root_len)
                .zip(project_root_path.iter().copied())
                .any(|(actual, expected)| !actual.eq_ignore_ascii_case(expected))
        {
            continue;
        }
        let storage_name = &path[root_len];
        let stream_name = &path[root_len + 1];
        let Some(frame_name) = stream_name.strip_prefix('\u{3}') else {
            continue;
        };
        if !frame_name.eq_ignore_ascii_case(VBFRAME_STREAM_PREFIX) {
            continue;
        }
        check_limit(
            "VBA designer count",
            frames.len().saturating_add(1),
            limits.max_modules,
        )?;
        let module = modules
            .iter()
            .find(|module| module.stream_name().eq_ignore_ascii_case(storage_name))
            .ok_or_else(|| {
                invalid(format!(
                    "designer storage {storage_name} has no matching dir module"
                ))
            })?;
        let components: Vec<&str> = path.iter().map(String::as_str).collect();
        let raw = read_limited_stream(ole, &components, limits.max_decompressed_stream_bytes)?;
        let frame = DesignerFrame::from_raw(module.name().to_owned(), raw, page, limits)?;
        frames
            .try_reserve(1)
            .map_err(|_| invalid("VBA designer frame allocation failed"))?;
        frames.push(frame);
    }
    Ok(frames)
}

fn parse_vbframe(
    expected_module_name: String,
    source: Text,
    limits: &Limits,
) -> Result<DesignerFrame, Error> {
    let lines = project_lines(source.text())?;
    if lines.len() < 3 || !lines[0].trim().eq_ignore_ascii_case("VERSION 5.00") {
        return Err(invalid("VBFrame stream is missing VERSION 5.00"));
    }
    let mut begin = lines[1].trim_matches([' ', '\t']);
    let Some((begin_keyword, remainder)) = take_vbframe_token(begin) else {
        return Err(invalid("VBFrame stream is missing its Begin record"));
    };
    if !begin_keyword.eq_ignore_ascii_case("Begin") {
        return Err(invalid("VBFrame stream is missing its Begin record"));
    }
    begin = remainder;
    let Some((class_id, remainder)) = take_vbframe_token(begin) else {
        return Err(invalid("VBFrame Begin record is missing its CLSID"));
    };
    let module_name = remainder.trim_matches([' ', '\t']);
    if module_name.is_empty() {
        return Err(invalid("VBFrame Begin record is missing its module name"));
    }
    if !valid_project_guid(class_id) || module_name != expected_module_name {
        return Err(invalid("VBFrame Begin record has invalid identity fields"));
    }
    let mut properties = Vec::new();
    let mut previous_property_index = None;
    let mut ended = false;
    for line in &lines[2..] {
        check_limit(
            "VBA designer property count",
            properties.len().saturating_add(1),
            limits.max_string_bytes,
        )?;
        let trimmed = line.trim_matches([' ', '\t']);
        if ended {
            return Err(invalid("VBFrame stream has data after End"));
        }
        if trimmed.eq_ignore_ascii_case("End") {
            ended = true;
            continue;
        }
        if trimmed.is_empty() || trimmed.starts_with('\'') {
            continue;
        }
        if let Some(property) = parse_designer_property(trimmed)? {
            let property_index = designer_property_index(property.name())
                .ok_or_else(|| invalid("VBFrame property is outside its vocabulary"))?;
            if previous_property_index.is_some_and(|previous| property_index <= previous) {
                return Err(invalid(
                    "VBFrame designer properties are out of order or duplicated",
                ));
            }
            previous_property_index = Some(property_index);
            properties
                .try_reserve(1)
                .map_err(|_| invalid("VBFrame property allocation failed"))?;
            properties.push(property);
        }
    }
    if !ended {
        return Err(invalid("VBFrame stream is missing its End record"));
    }
    Ok(DesignerFrame {
        module_name: expected_module_name,
        class_id: class_id.to_owned(),
        properties,
        source,
    })
}

fn take_vbframe_token(value: &str) -> Option<(&str, &str)> {
    let value = value.trim_start_matches([' ', '\t']);
    let split = value.find([' ', '\t'])?;
    Some((&value[..split], &value[split..]))
}

fn parse_designer_property(line: &str) -> Result<Option<DesignerProperty>, Error> {
    let (value_line, comment) = split_designer_comment(line);
    let Some((name, value)) = value_line.split_once('=') else {
        return Ok(None);
    };
    let name = name.trim();
    let value = value.trim();
    let canonical_name = DESIGNER_PROPERTY_NAMES
        .iter()
        .copied()
        .find(|candidate| name.eq_ignore_ascii_case(candidate));
    let Some(canonical_name) = canonical_name else {
        return Ok(None);
    };
    match canonical_name {
        "Caption" | "Tag" => {
            let text = parse_quoted_value(value, "VBFrame text property")?;
            if text.chars().count() > 130 {
                return Err(invalid("VBFrame text property exceeds 130 characters"));
            }
        },
        "ClientHeight" | "ClientLeft" | "ClientTop" | "ClientWidth" => {
            let number = value
                .parse::<f64>()
                .map_err(|_| invalid("VBFrame coordinate is not a FLOAT"))?;
            if !number.is_finite() {
                return Err(invalid("VBFrame coordinate is not finite"));
            }
        },
        "Enabled" | "RightToLeft" | "ShowModal" | "Visible" | "WhatsThisButton"
        | "WhatsThisHelp" => {
            if !matches!(value, "0" | "-1") {
                return Err(invalid("VBFrame Boolean property is not 0 or -1"));
            }
        },
        "HelpContextID" | "TypeInfoVer" => {
            value
                .parse::<i32>()
                .map_err(|_| invalid("VBFrame integer property is not an INT32"))?;
        },
        "StartUpPosition" => {
            if !matches!(value, "0" | "1" | "2" | "3") {
                return Err(invalid("VBFrame StartUpPosition is outside 0..=3"));
            }
        },
        _ => unreachable!(),
    }
    Ok(Some(DesignerProperty::new(
        canonical_name,
        value,
        comment.map(str::to_owned),
    )))
}

const DESIGNER_PROPERTY_NAMES: [&str; 15] = [
    "Caption",
    "ClientHeight",
    "ClientLeft",
    "ClientTop",
    "ClientWidth",
    "Enabled",
    "HelpContextID",
    "RightToLeft",
    "ShowModal",
    "StartUpPosition",
    "Tag",
    "TypeInfoVer",
    "Visible",
    "WhatsThisButton",
    "WhatsThisHelp",
];

fn designer_property_index(name: &str) -> Option<usize> {
    DESIGNER_PROPERTY_NAMES
        .iter()
        .position(|candidate| candidate.eq_ignore_ascii_case(name))
}

fn split_designer_comment(line: &str) -> (&str, Option<&str>) {
    let bytes = line.as_bytes();
    let mut quoted = false;
    let mut index = 0usize;
    while index < bytes.len() {
        match bytes[index] {
            b'"' => {
                if quoted && bytes.get(index + 1) == Some(&b'"') {
                    index += 2;
                    continue;
                }
                quoted = !quoted;
            },
            b'\'' if !quoted => {
                return (&line[..index], Some(&line[index + 1..]));
            },
            _ => {},
        }
        index += 1;
    }
    (line, None)
}

fn parse_project_lk(bytes: &[u8], limits: &Limits) -> Result<Vec<LicenseInfo>, Error> {
    let version = bytes
        .get(..2)
        .map(|raw| u16::from_le_bytes([raw[0], raw[1]]))
        .ok_or_else(|| invalid("PROJECTlk stream is missing its version"))?;
    if version != 1 {
        return Err(invalid(format!(
            "PROJECTlk stream version {version} is not supported"
        )));
    }
    let count = bytes
        .get(2..6)
        .map(|raw| u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]))
        .ok_or_else(|| invalid("PROJECTlk stream is missing its count"))?;
    let count = usize::try_from(count)
        .map_err(|_| invalid("PROJECTlk license count does not fit usize"))?;
    check_limit("VBA license count", count, limits.max_licenses)?;
    const LICENSEINFO_MINIMUM_BYTES: usize = 16 + 4 + 4;
    let minimum_bytes = count
        .checked_mul(LICENSEINFO_MINIMUM_BYTES)
        .and_then(|value| value.checked_add(6))
        .ok_or_else(|| invalid("PROJECTlk minimum size overflows usize"))?;
    if bytes.len() < minimum_bytes {
        return Err(invalid(format!(
            "PROJECTlk stream is truncated before its {count} LICENSEINFO records"
        )));
    }
    let mut cursor = 6usize;
    let mut entries = Vec::new();
    entries
        .try_reserve(count)
        .map_err(|_| invalid("PROJECTlk license allocation failed"))?;
    for index in 0..count {
        let class_id = bytes
            .get(cursor..cursor.saturating_add(16))
            .ok_or_else(|| invalid(format!("PROJECTlk LICENSEINFO {index} is truncated")))?;
        let mut class_id_array = [0u8; 16];
        class_id_array.copy_from_slice(class_id);
        cursor = cursor
            .checked_add(16)
            .ok_or_else(|| invalid("PROJECTlk cursor overflow"))?;

        let key_length = bytes
            .get(cursor..cursor.saturating_add(4))
            .map(|raw| u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]))
            .ok_or_else(|| invalid(format!("PROJECTlk LICENSEINFO {index} is truncated")))?;
        cursor = cursor
            .checked_add(4)
            .ok_or_else(|| invalid("PROJECTlk cursor overflow"))?;
        let key_length = usize::try_from(key_length)
            .map_err(|_| invalid("PROJECTlk license key length does not fit usize"))?;
        check_limit("VBA license key bytes", key_length, limits.max_string_bytes)?;
        let key_end = cursor
            .checked_add(key_length)
            .ok_or_else(|| invalid("PROJECTlk license key length overflows stream"))?;
        let license_key = bytes
            .get(cursor..key_end)
            .ok_or_else(|| invalid(format!("PROJECTlk LICENSEINFO {index} key is truncated")))?
            .to_vec();
        cursor = key_end;

        let license_required = bytes
            .get(cursor..cursor.saturating_add(4))
            .map(|raw| u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]))
            .ok_or_else(|| invalid(format!("PROJECTlk LICENSEINFO {index} is truncated")))?;
        cursor = cursor
            .checked_add(4)
            .ok_or_else(|| invalid("PROJECTlk cursor overflow"))?;
        entries.push(LicenseInfo::from_raw(
            class_id_array,
            license_key,
            license_required,
        ));
    }
    if cursor != bytes.len() {
        return Err(invalid("PROJECTlk stream has trailing bytes"));
    }
    Ok(entries)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ProjectTextSection {
    Properties,
    HostExtenders,
    Workspace,
}

fn project_lines(text: &str) -> Result<Vec<&str>, Error> {
    if text.is_empty() {
        return Err(invalid("VBA text stream is empty"));
    }
    let bytes = text.as_bytes();
    let mut lines = Vec::new();
    let mut start = 0usize;
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        if !matches!(bytes[cursor], b'\r' | b'\n') {
            cursor += 1;
            continue;
        }
        if cursor + 1 >= bytes.len()
            || !((bytes[cursor] == b'\r' && bytes[cursor + 1] == b'\n')
                || (bytes[cursor] == b'\n' && bytes[cursor + 1] == b'\r'))
        {
            return Err(invalid("PROJECT stream contains a newline other than NWLN"));
        }
        lines
            .try_reserve(1)
            .map_err(|_| invalid("PROJECT line allocation failed"))?;
        lines.push(&text[start..cursor]);
        cursor += 2;
        start = cursor;
    }
    if start < text.len() {
        return Err(invalid("VBA text stream is missing its final NWLN"));
    }
    Ok(lines)
}

fn parse_project_property_line(line: &str, limits: &Limits) -> Result<ProjectTextRecord, Error> {
    let Some((key, value)) = line.split_once('=') else {
        return Ok(ProjectTextRecord::Unknown(line.to_owned()));
    };
    let key = key.trim_matches([' ', '\t']);
    let value = value.trim_matches([' ', '\t']);
    if key.eq_ignore_ascii_case("ID") {
        let value = parse_quoted_value(value, "PROJECT ID")?;
        if !valid_project_guid(&value) {
            return Err(invalid("PROJECT ID is not a GUID"));
        }
        return Ok(ProjectTextRecord::Id(value));
    }
    if key.eq_ignore_ascii_case("Document") {
        let Some((name, version)) = value.split_once('/') else {
            return Err(invalid("PROJECT Document record is missing its version"));
        };
        validate_project_module_identifier(name)?;
        let type_library_version = parse_hex_int32(version, "PROJECT Document version")?;
        return Ok(ProjectTextRecord::Module {
            name: name.to_owned(),
            kind: ProjectModuleKind::Document {
                type_library_version,
            },
        });
    }
    if key.eq_ignore_ascii_case("Module")
        || key.eq_ignore_ascii_case("Class")
        || key.eq_ignore_ascii_case("BaseClass")
    {
        validate_project_module_identifier(value)?;
        let kind = if key.eq_ignore_ascii_case("Module") {
            ProjectModuleKind::Standard
        } else if key.eq_ignore_ascii_case("Class") {
            ProjectModuleKind::Class
        } else {
            ProjectModuleKind::Designer
        };
        return Ok(ProjectTextRecord::Module {
            name: value.to_owned(),
            kind,
        });
    }
    if key.eq_ignore_ascii_case("Package") {
        if !valid_project_guid(value) {
            return Err(invalid("PROJECT Package record is not a GUID"));
        }
        return Ok(ProjectTextRecord::Package(value.to_owned()));
    }
    if key.eq_ignore_ascii_case("HelpFile") {
        return Ok(ProjectTextRecord::HelpFile(parse_path(value, limits)?));
    }
    if key.eq_ignore_ascii_case("ExeName32") {
        return Ok(ProjectTextRecord::ExeName32(parse_path(value, limits)?));
    }
    if key.eq_ignore_ascii_case("Name") {
        let name = parse_quoted_value(value, "PROJECT Name")?;
        if name.is_empty() || name.chars().count() > 128 {
            return Err(invalid(
                "PROJECT Name is outside its 1..=128 character bound",
            ));
        }
        return Ok(ProjectTextRecord::Name(name));
    }
    if key.eq_ignore_ascii_case("HelpContextID") {
        return Ok(ProjectTextRecord::HelpContextId(
            parse_quoted_value(value, "PROJECT HelpContextID")?
                .parse::<i32>()
                .map_err(|_| invalid("PROJECT HelpContextID is not an INT32"))?,
        ));
    }
    if key.eq_ignore_ascii_case("Description") {
        let description = parse_quoted_value(value, "PROJECT Description")?;
        if description.chars().count() > 2000 {
            return Err(invalid("PROJECT Description exceeds 2000 characters"));
        }
        return Ok(ProjectTextRecord::Description(description));
    }
    if key.eq_ignore_ascii_case("VersionCompatible32") {
        let version = parse_quoted_value(value, "PROJECT VersionCompatible32")?;
        if version != "393222000" {
            return Err(invalid("PROJECT VersionCompatible32 is not 393222000"));
        }
        return Ok(ProjectTextRecord::VersionCompatible32(version));
    }
    if key.eq_ignore_ascii_case("CMG") {
        return Ok(ProjectTextRecord::ProtectionState(parse_hex_string(
            value,
            22,
            28,
            "PROJECT CMG",
        )?));
    }
    if key.eq_ignore_ascii_case("DPB") {
        return Ok(ProjectTextRecord::Password(parse_hex_string(
            value,
            16,
            usize::MAX,
            "PROJECT DPB",
        )?));
    }
    if key.eq_ignore_ascii_case("GC") {
        return Ok(ProjectTextRecord::VisibilityState(parse_hex_string(
            value,
            16,
            22,
            "PROJECT GC",
        )?));
    }
    Ok(ProjectTextRecord::Unknown(line.to_owned()))
}

fn parse_host_extender_line(line: &str) -> Result<ProjectTextRecord, Error> {
    let Some((index, rest)) = line.split_once('=') else {
        return Ok(ProjectTextRecord::Unknown(line.to_owned()));
    };
    let index = index.trim_matches([' ', '\t']);
    let rest = rest.trim_matches([' ', '\t']);
    let mut fields = rest.split(';');
    let Some(guid) = fields.next() else {
        return Err(invalid("PROJECT host extender is missing its GUID"));
    };
    let Some(lib_name) = fields.next() else {
        return Err(invalid("PROJECT host extender is missing its library name"));
    };
    let Some(creation_flags) = fields.next() else {
        return Err(invalid(
            "PROJECT host extender is missing its creation flags",
        ));
    };
    if fields.next().is_some() || !valid_project_guid(guid) {
        return Err(invalid(
            "PROJECT host extender GUID or field count is invalid",
        ));
    }
    if lib_name
        .bytes()
        .any(|byte| !(0x21..=0x3a).contains(&byte) && !(0x3c..=0xff).contains(&byte))
    {
        return Err(invalid("PROJECT host extender library name is invalid"));
    }
    let index = parse_hex_int32(index, "PROJECT host extender index")?;
    let creation_flags = parse_hex_int32(creation_flags, "PROJECT host extender flags")?;
    Ok(ProjectTextRecord::HostExtender(ProjectHostExtender::new(
        index,
        guid,
        lib_name,
        creation_flags,
    )))
}

fn parse_workspace_line(line: &str) -> Result<ProjectTextRecord, Error> {
    let Some((module_name, state)) = line.split_once('=') else {
        return Ok(ProjectTextRecord::Unknown(line.to_owned()));
    };
    let module_name = module_name.trim_matches([' ', '\t']);
    let state = state.trim_matches([' ', '\t']);
    validate_project_module_identifier(module_name)?;
    if !state.contains(',') {
        return Ok(ProjectTextRecord::Unknown(line.to_owned()));
    }
    let fields: Vec<_> = state.split(',').map(str::trim).collect();
    let code = parse_window_state_fields(fields.get(..5).unwrap_or_default())?;
    let designer = match fields.len() {
        5 => None,
        10 => Some(parse_window_state_fields(&fields[5..])?),
        _ => return Err(invalid("PROJECT window has an invalid field count")),
    };
    Ok(ProjectTextRecord::Window(ProjectWindow::new(
        module_name,
        code,
        designer,
    )))
}

fn parse_window_state_fields(fields: &[&str]) -> Result<ProjectWindowState, Error> {
    let [left, top, right, bottom, state] = fields else {
        return Err(invalid("PROJECT window is missing fields"));
    };
    let left = parse_int32_field(Some(left), "PROJECT window left")?;
    let top = parse_int32_field(Some(top), "PROJECT window top")?;
    let right = parse_int32_field(Some(right), "PROJECT window right")?;
    let bottom = parse_int32_field(Some(bottom), "PROJECT window bottom")?;
    let mut seen = 0u8;
    if state.len() > 3 {
        return Err(invalid("PROJECT window state is invalid"));
    }
    for byte in state.bytes() {
        let bit = match byte.to_ascii_uppercase() {
            b'C' => 1,
            b'Z' => 2,
            b'I' => 4,
            _ => return Err(invalid("PROJECT window state is invalid")),
        };
        if seen & bit != 0 {
            return Err(invalid("PROJECT window state is invalid"));
        }
        seen |= bit;
    }
    Ok(ProjectWindowState::new(left, top, right, bottom, *state))
}

fn parse_int32_field(value: Option<&str>, field: &'static str) -> Result<i32, Error> {
    value
        .map(str::trim)
        .ok_or_else(|| invalid(format!("{field} is missing")))?
        .parse::<i32>()
        .map_err(|_| invalid(format!("{field} is not an INT32")))
}

fn parse_hex_int32(value: &str, field: &'static str) -> Result<u32, Error> {
    let Some(value) = value
        .strip_prefix("&H")
        .or_else(|| value.strip_prefix("&h"))
    else {
        return Err(invalid(format!("{field} is not a HEXINT32")));
    };
    if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid(format!("{field} is not a HEXINT32")));
    }
    let value = u32::from_str_radix(value, 16)
        .map_err(|_| invalid(format!("{field} is not a HEXINT32")))?;
    if value > i32::MAX as u32 {
        return Err(invalid(format!("{field} is not a signed HEXINT32")));
    }
    Ok(value)
}

fn parse_hex_string(
    value: &str,
    minimum: usize,
    maximum: usize,
    field: &'static str,
) -> Result<String, Error> {
    let value = parse_quoted_value(value, field)?;
    if !(minimum..=maximum).contains(&value.len())
        || !value.bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        return Err(invalid(format!(
            "{field} is not a bounded hexadecimal value"
        )));
    }
    Ok(value)
}

fn parse_path(value: &str, limits: &Limits) -> Result<String, Error> {
    let path = parse_quoted_value(value, "PROJECT path")?;
    if path.chars().count() > 259 {
        return Err(invalid("PROJECT path exceeds 259 characters"));
    }
    check_limit("PROJECT path bytes", path.len(), limits.max_string_bytes)?;
    Ok(path)
}

fn parse_quoted_value(value: &str, field: &'static str) -> Result<String, Error> {
    let Some(inner) = value
        .strip_prefix('"')
        .and_then(|value| value.strip_suffix('"'))
    else {
        return Err(invalid(format!("{field} is not quoted")));
    };
    let mut output = String::with_capacity(inner.len());
    let mut chars = inner.chars().peekable();
    while let Some(character) = chars.next() {
        if character == '"' {
            if chars.next_if_eq(&'"').is_none() {
                return Err(invalid(format!("{field} contains an unescaped quote")));
            }
        } else if !valid_quoted_character(character) {
            return Err(invalid(format!("{field} contains a control character")));
        }
        output.push(character);
    }
    Ok(output)
}

fn valid_quoted_character(character: char) -> bool {
    character == '\t' || !character.is_control() || ('\u{7f}'..='\u{ff}').contains(&character)
}

fn valid_project_guid(value: &str) -> bool {
    let bytes = value.as_bytes();
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

fn validate_project_module_identifier(value: &str) -> Result<(), Error> {
    if value.is_empty()
        || value.chars().count() > 31
        || value
            .chars()
            .any(|character| matches!(character, '\0' | '\r' | '\n'))
    {
        return Err(invalid(
            "PROJECT module identifier is outside its 1..=31 character bound",
        ));
    }
    Ok(())
}

fn render_project_text_record(record: &ProjectTextRecord, output: &mut String) {
    match record {
        ProjectTextRecord::Id(value) => {
            let _ = writeln!(output, "ID=\"{}\"", quote_project_value(value));
        },
        ProjectTextRecord::Module { name, kind } => match kind {
            ProjectModuleKind::Standard => {
                let _ = writeln!(output, "Module={name}");
            },
            ProjectModuleKind::Class => {
                let _ = writeln!(output, "Class={name}");
            },
            ProjectModuleKind::Designer => {
                let _ = writeln!(output, "BaseClass={name}");
            },
            ProjectModuleKind::Document {
                type_library_version,
            } => {
                let _ = writeln!(output, "Document={name}/&H{type_library_version:08X}");
            },
        },
        ProjectTextRecord::Package(value) => {
            let _ = writeln!(output, "Package={value}");
        },
        ProjectTextRecord::HelpFile(value) => {
            let _ = writeln!(output, "HelpFile=\"{}\"", quote_project_value(value));
        },
        ProjectTextRecord::ExeName32(value) => {
            let _ = writeln!(output, "ExeName32=\"{}\"", quote_project_value(value));
        },
        ProjectTextRecord::Name(value) => {
            let _ = writeln!(output, "Name=\"{}\"", quote_project_value(value));
        },
        ProjectTextRecord::HelpContextId(value) => {
            let _ = writeln!(output, "HelpContextID=\"{value}\"");
        },
        ProjectTextRecord::Description(value) => {
            let _ = writeln!(output, "Description=\"{}\"", quote_project_value(value));
        },
        ProjectTextRecord::VersionCompatible32(value) => {
            let _ = writeln!(output, "VersionCompatible32=\"{value}\"");
        },
        ProjectTextRecord::ProtectionState(value) => {
            let _ = writeln!(output, "CMG=\"{value}\"");
        },
        ProjectTextRecord::Password(value) => {
            let _ = writeln!(output, "DPB=\"{value}\"");
        },
        ProjectTextRecord::VisibilityState(value) => {
            let _ = writeln!(output, "GC=\"{value}\"");
        },
        ProjectTextRecord::HostExtenderSection => output.push_str("[Host Extender Info]\r\n"),
        ProjectTextRecord::HostExtender(value) => {
            let _ = writeln!(
                output,
                "&H{:08X}={};{};&H{:08X}",
                value.index, value.guid, value.lib_name, value.creation_flags
            );
        },
        ProjectTextRecord::WorkspaceSection => output.push_str("[Workspace]\r\n"),
        ProjectTextRecord::Window(value) => {
            let _ = write!(output, "{}=", value.module_name);
            render_window_state(&value.code, output);
            if let Some(designer) = &value.designer {
                output.push_str(", ");
                render_window_state(designer, output);
            }
            output.push_str("\r\n");
        },
        ProjectTextRecord::Blank => output.push_str("\r\n"),
        ProjectTextRecord::Unknown(value) => {
            output.push_str(value);
            output.push_str("\r\n");
        },
    }
}

fn render_window_state(value: &ProjectWindowState, output: &mut String) {
    let _ = write!(
        output,
        "{}, {}, {}, {}, {}",
        value.left, value.top, value.right, value.bottom, value.state
    );
}

fn quote_project_value(value: &str) -> String {
    value.replace('"', "\"\"")
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test fixtures and assertions panic intentionally on failure"
    )]

    use super::*;
    use litchi_cfb::OleWriter;
    use std::io::Cursor;

    fn push_record(bytes: &mut Vec<u8>, id: u16, value: &[u8]) {
        bytes.extend_from_slice(&id.to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(value.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(value);
    }

    fn push_pair(bytes: &mut Vec<u8>, id: u16, value: &str, reserved: u16) {
        push_record(bytes, id, value.as_bytes());
        bytes.extend_from_slice(&reserved.to_le_bytes());
        let unicode: Vec<u8> = value.encode_utf16().flat_map(u16::to_le_bytes).collect();
        bytes.extend_from_slice(&u32::try_from(unicode.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(&unicode);
    }

    fn push_project_information(bytes: &mut Vec<u8>, name: &str) {
        push_record(bytes, 0x0001, &1u32.to_le_bytes());
        push_record(bytes, 0x0002, &0x0409u32.to_le_bytes());
        push_record(bytes, 0x0014, &0x0409u32.to_le_bytes());
        push_record(bytes, 0x0003, &1252u16.to_le_bytes());
        push_record(bytes, 0x0004, name.as_bytes());
        push_pair(bytes, 0x0005, "", 0x0040);
        push_record(bytes, 0x0006, &[]);
        bytes.extend_from_slice(&0x003du16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        push_record(bytes, 0x0007, &0u32.to_le_bytes());
        push_record(bytes, 0x0008, &0u32.to_le_bytes());
        bytes.extend_from_slice(&0x0009u16.to_le_bytes());
        bytes.extend_from_slice(&4u32.to_le_bytes());
        bytes.extend_from_slice(&1u32.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
    }

    fn literal_container(data: &[u8]) -> Vec<u8> {
        let mut encoded = vec![0x01];
        for decompressed_chunk in data.chunks(3_000) {
            let mut chunk = Vec::with_capacity(decompressed_chunk.len() + 375);
            for literals in decompressed_chunk.chunks(8) {
                chunk.push(0);
                chunk.extend_from_slice(literals);
            }
            let header = 0xb000 | u16::try_from(chunk.len() - 1).unwrap();
            encoded.extend_from_slice(&header.to_le_bytes());
            encoded.extend_from_slice(&chunk);
        }
        encoded
    }

    fn sample_dir() -> Vec<u8> {
        let mut bytes = Vec::new();
        push_project_information(&mut bytes, "Sample");
        push_record(&mut bytes, 0x000f, &1u16.to_le_bytes());
        push_record(&mut bytes, 0x0013, &0xffffu16.to_le_bytes());
        push_record(&mut bytes, 0x0019, b"Module1");
        push_pair(&mut bytes, 0x001a, "Module1", 0x0032);
        push_pair(&mut bytes, 0x001c, "", 0x0048);
        push_record(&mut bytes, 0x0031, &3u32.to_le_bytes());
        push_record(&mut bytes, 0x001e, &0u32.to_le_bytes());
        push_record(&mut bytes, 0x002c, &0xffffu16.to_le_bytes());
        bytes.extend_from_slice(&0x0021u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0x002bu16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0x0010u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        literal_container(&bytes)
    }

    fn sample_project_bytes(source: &[u8]) -> Vec<u8> {
        let mut module_stream = vec![7, 8, 9];
        module_stream.extend_from_slice(&literal_container(source));

        let mut writer = OleWriter::new();
        writer
            .create_stream(&["PROJECT"], b"ID=\"Sample\"\r\nModule=Module1\r\n")
            .expect("PROJECT stream");
        writer
            .create_stream(&["VBA", "_VBA_PROJECT"], &[0; 8])
            .expect("version stream");
        writer
            .create_stream(&["VBA", "dir"], &sample_dir())
            .expect("dir stream");
        writer
            .create_stream(&["VBA", "Module1"], &module_stream)
            .expect("module stream");
        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).expect("standalone CFB");
        cursor.into_inner()
    }

    #[test]
    fn parses_project_lk_license_records_with_raw_boolean_values() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&1u16.to_le_bytes());
        bytes.extend_from_slice(&1u32.to_le_bytes());
        bytes.extend_from_slice(&[0xabu8; 16]);
        bytes.extend_from_slice(&3u32.to_le_bytes());
        bytes.extend_from_slice(&[1, 2, 3]);
        bytes.extend_from_slice(&0xfeed_beefu32.to_le_bytes());

        let entries = parse_project_lk(&bytes, &Limits::default()).unwrap();
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].class_id(), [0xab; 16]);
        assert_eq!(entries[0].license_key(), [1, 2, 3]);
        assert_eq!(entries[0].license_required_raw(), 0xfeed_beef);
        assert!(entries[0].license_required());
    }

    #[test]
    fn rejects_project_lk_version_truncation_trailing_and_count_limits() {
        let mut bad_version = vec![2, 0, 0, 0, 0, 0];
        assert!(matches!(
            parse_project_lk(&bad_version, &Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("version")
        ));

        bad_version[0] = 1;
        bad_version[2] = 1;
        assert!(matches!(
            parse_project_lk(&bad_version, &Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("truncated")
        ));

        let mut trailing = vec![1, 0, 0, 0, 0, 0, 0xaa];
        assert!(matches!(
            parse_project_lk(&trailing, &Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("trailing")
        ));

        trailing[2..6].copy_from_slice(&1u32.to_le_bytes());
        let limits = Limits {
            max_licenses: 0,
            ..Limits::default()
        };
        assert!(matches!(
            parse_project_lk(&trailing, &limits),
            Err(Error::LimitExceeded {
                limit: "VBA license count",
                actual: 1,
                maximum: 0,
            })
        ));
    }

    #[test]
    fn parses_and_edits_typed_project_text_records() {
        let source = concat!(
            "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n",
            "Document=ThisDocument/&H00010000\r\n",
            "Module=Module1\r\n",
            "Class=Class1\r\n",
            "BaseClass=Form1\r\n",
            "Package={00000000-0000-0000-0000-000000000001}\r\n",
            "HelpFile=\"C:\\\\help.chm\"\r\n",
            "ExeName32=\"host.exe\"\r\n",
            "Name=\"Sample\"\r\n",
            "HelpContextID=\"-7\"\r\n",
            "Description=\"A \"\"quoted\"\" description\"\r\n",
            "VersionCompatible32=\"393222000\"\r\n",
            "CMG=\"0000000000000000000000\"\r\n",
            "DPB=\"0000000000000000\"\r\n",
            "GC=\"0000000000000000\"\r\n",
            "\r\n[Host Extender Info]\r\n",
            "&H00000001={00000000-0000-0000-0000-000000000002};VBE;&H00000001\r\n",
            "\r\n[Workspace]\r\n",
            "Module1=1, 2, 3, 4, CZ, 5, 6, 7, 8, I\r\n",
            "VendorExtension=keep\r\n",
        );
        let mut project = ProjectText::parse(source, &Limits::default()).unwrap();
        assert_eq!(project.project_name(), Some("Sample"));
        assert_eq!(project.records().len(), 22);
        assert!(project.records().iter().any(|record| matches!(
            record,
            ProjectTextRecord::Module {
                kind: ProjectModuleKind::Document { .. },
                ..
            }
        )));
        assert!(
            project
                .records()
                .iter()
                .any(|record| matches!(record, ProjectTextRecord::HostExtender(_)))
        );
        assert!(
            project
                .records()
                .iter()
                .any(|record| matches!(record, ProjectTextRecord::Window(_)))
        );
        assert!(
            project
                .records()
                .iter()
                .any(|record| matches!(record, ProjectTextRecord::Unknown(_)))
        );
        assert_eq!(project.to_text(), source);

        project.rename_module("Module1", "Renamed").unwrap();
        project.set_project_name("Changed").unwrap();
        let edited = project.to_text();
        assert!(edited.contains("Name=\"Changed\"\r\n"));
        assert!(edited.contains("Module=Renamed\r\n"));
        assert!(edited.contains("Renamed=1, 2, 3, 4, CZ, 5, 6, 7, 8, I\r\n"));
        assert!(edited.contains("VendorExtension=keep\r\n"));
        assert!(
            project
                .encode(Mbcs::WINDOWS_1252, &Limits::default())
                .is_ok()
        );
    }

    #[test]
    fn project_text_noops_preserve_exact_source_and_grammar_is_case_insensitive() {
        let source = concat!(
            "id=\"{00000000-0000-0000-0000-000000000000}\"\r\n",
            "document=ThisDocument/&h00010000\r\n",
            "baseclass=Form One\r\n",
            "name=\"Sample\"\r\n",
            "versioncompatible32=\"393222000\"\r\n",
            "\r\n[workspace]\r\n",
            "Form One=1, 2, 3, 4, czi\r\n",
        );
        let mut project = ProjectText::parse(source, &Limits::default()).unwrap();
        project.set_project_name("Sample").unwrap();
        project.rename_module("Form One", "Form One").unwrap();
        assert_eq!(project.to_text(), source);
        assert_eq!(
            project
                .encode(Mbcs::WINDOWS_1252, &Limits::default())
                .unwrap(),
            source.as_bytes()
        );
        assert!(project.records().iter().any(|record| matches!(
            record,
            ProjectTextRecord::Module {
                kind: ProjectModuleKind::Designer,
                ..
            }
        )));
        let window = project
            .records()
            .iter()
            .find_map(|record| match record {
                ProjectTextRecord::Window(window) => Some(window),
                _ => None,
            })
            .unwrap();
        assert_eq!(window.code().state(), "czi");
    }

    #[test]
    fn project_text_rejects_duplicate_window_flags_and_unsigned_hexint32() {
        for state in ["CIC", "CZIC", "Q"] {
            let source =
                format!("Name=\"Sample\"\r\n\r\n[Workspace]\r\nModule1=1, 2, 3, 4, {state}\r\n");
            assert!(ProjectText::parse(&source, &Limits::default()).is_err());
        }
        let source = "Document=ThisDocument/&HFFFFFFFF\r\n";
        assert!(matches!(
            ProjectText::parse(source, &Limits::default()),
            Err(Error::InvalidData(message)) if message.contains("signed HEXINT32")
        ));
    }

    #[test]
    fn cached_and_exact_noop_paths_recheck_limits() {
        let bytes = crate::build::Project::new("Sample")
            .module(crate::build::Module::standard(
                "Module1",
                "Attribute VB_Name = \"Module1\"\r\n",
            ))
            .finish(&Limits::default())
            .unwrap()
            .into_bytes();
        let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let project = Project::open(&mut ole, &[], &Limits::default()).unwrap();
        let permissive = Limits::default();
        let text = project.project_text(&permissive).unwrap();
        assert_eq!(
            project.encode_project_text(text, &permissive).unwrap(),
            project.project_properties().raw()
        );

        let tiny_string = Limits {
            max_string_bytes: 1,
            ..permissive
        };
        assert!(matches!(
            project.project_text(&tiny_string),
            Err(Error::LimitExceeded {
                limit: "PROJECT text line bytes",
                ..
            })
        ));
        assert!(matches!(
            project.encode_project_text(text, &tiny_string),
            Err(Error::LimitExceeded {
                limit: "PROJECT text line bytes",
                ..
            })
        ));
        assert!(matches!(
            project.encode_dir_with_references(project.references(), &tiny_string),
            Err(Error::LimitExceeded {
                limit: "VBA string bytes",
                ..
            })
        ));

        let no_modules = Limits {
            max_modules: 0,
            ..permissive
        };
        assert!(matches!(
            project.project_text(&no_modules),
            Err(Error::LimitExceeded {
                limit: "VBA module count",
                actual: 1,
                maximum: 0,
            })
        ));
        assert!(matches!(
            project.project_wm(&no_modules),
            Err(Error::LimitExceeded {
                limit: "VBA module count",
                actual: 1,
                maximum: 0,
            })
        ));
        let maps = project.project_wm(&permissive).unwrap().unwrap().to_vec();
        assert!(matches!(
            project.encode_project_wm(&maps, &no_modules),
            Err(Error::LimitExceeded {
                limit: "VBA module count",
                actual: 1,
                maximum: 0,
            })
        ));
    }

    #[test]
    fn parses_bounded_vbframe_properties_without_interpreting_code() {
        let raw = concat!(
            "VERSION 5.00\r\n",
            "Begin {00000000-0000-0000-0000-000000000000} Form1\r\n",
            " Caption=\"Hello\" ' preserved comment\r\n",
            " ClientHeight=120.5\r\n",
            " Enabled=-1\r\n",
            " StartUpPosition=2\r\n",
            "End\r\n",
        )
        .as_bytes()
        .to_vec();
        let frame =
            DesignerFrame::from_raw("Form1", raw, Mbcs::WINDOWS_1252, &Limits::default()).unwrap();
        assert_eq!(frame.module_name(), "Form1");
        assert_eq!(frame.properties().len(), 4);
        assert_eq!(frame.properties()[0].name(), "Caption");
        assert_eq!(frame.properties()[0].comment(), Some(" preserved comment"));
        assert_eq!(frame.properties()[3].value(), "2");

        let invalid = b"VERSION 5.00\r\nBegin {00000000-0000-0000-0000-000000000000} Form1\r\nEnabled=1\r\nEnd\r\n";
        assert!(matches!(
            DesignerFrame::from_raw(
                "Form1",
                invalid.to_vec(),
                Mbcs::WINDOWS_1252,
                &Limits::default(),
            ),
            Err(Error::InvalidData(message)) if message.contains("Boolean")
        ));
        let limits = Limits {
            max_decompressed_stream_bytes: 4,
            ..Limits::default()
        };
        assert!(matches!(
            DesignerFrame::from_raw(
                "Form1",
                b"VERSION 5.00\r\n".to_vec(),
                Mbcs::WINDOWS_1252,
                &limits,
            ),
            Err(Error::LimitExceeded {
                limit: "VBA designer stream bytes",
                ..
            })
        ));

        let spaced = concat!(
            "version 5.00\r\n",
            "begin {00000000-0000-0000-0000-000000000000} Form One\r\n",
            " caption=\"Hello\"\r\n",
            " enabled=-1\r\n",
            "end\r\n",
        );
        let spaced = DesignerFrame::from_raw(
            "Form One",
            spaced.as_bytes().to_vec(),
            Mbcs::WINDOWS_1252,
            &Limits::default(),
        )
        .unwrap();
        assert_eq!(spaced.module_name(), "Form One");
        assert_eq!(spaced.properties()[0].name(), "Caption");

        let out_of_order = concat!(
            "VERSION 5.00\r\n",
            "Begin {00000000-0000-0000-0000-000000000000} Form1\r\n",
            "Enabled=-1\r\n",
            "Caption=\"Hello\"\r\n",
            "End\r\n",
        );
        assert!(matches!(
            DesignerFrame::from_raw(
                "Form1",
                out_of_order.as_bytes().to_vec(),
                Mbcs::WINDOWS_1252,
                &Limits::default(),
            ),
            Err(Error::InvalidData(message)) if message.contains("out of order")
        ));

        let duplicate = concat!(
            "VERSION 5.00\r\n",
            "Begin {00000000-0000-0000-0000-000000000000} Form1\r\n",
            "Caption=\"Hello\"\r\n",
            "caption=\"Again\"\r\n",
            "End\r\n",
        );
        assert!(matches!(
            DesignerFrame::from_raw(
                "Form1",
                duplicate.as_bytes().to_vec(),
                Mbcs::WINDOWS_1252,
                &Limits::default(),
            ),
            Err(Error::InvalidData(message)) if message.contains("duplicated")
        ));

        let trailing_blank = concat!(
            "VERSION 5.00\r\n",
            "Begin {00000000-0000-0000-0000-000000000000} Form1\r\n",
            "End\r\n",
            "\r\n",
        );
        assert!(matches!(
            DesignerFrame::from_raw(
                "Form1",
                trailing_blank.as_bytes().to_vec(),
                Mbcs::WINDOWS_1252,
                &Limits::default(),
            ),
            Err(Error::InvalidData(message)) if message.contains("after End")
        ));
    }

    #[test]
    fn reads_borrowed_standalone_project_with_a_preparse_size_bound() {
        let source = b"Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nEnd Sub\r\n";
        let bytes = sample_project_bytes(source);
        let project = Project::read(&bytes).expect("borrowed project");

        assert_eq!(project.name(), "Sample");
        assert_eq!(project.modules()[0].source().raw(), source);

        let limits = Limits {
            max_cfb_bytes: bytes.len().saturating_sub(1),
            ..Limits::default()
        };
        assert!(matches!(
            Project::read_with(&bytes, &limits),
            Err(Error::LimitExceeded {
                limit: "standalone VBA CFB bytes",
                actual,
                maximum,
            }) if actual == bytes.len() && maximum == limits.max_cfb_bytes
        ));

        let limits = Limits {
            max_total_source_bytes: 1,
            ..Limits::default()
        };
        assert!(matches!(
            Project::read_with(&bytes, &limits),
            Err(Error::LimitExceeded {
                limit: "aggregate VBA module source bytes",
                maximum: 1,
                ..
            })
        ));
    }

    #[test]
    fn borrowed_standalone_reader_returns_typed_cfb_errors() {
        assert!(matches!(
            Project::read(b"not a compound file"),
            Err(Error::Cfb(_))
        ));
    }

    #[test]
    fn opens_project_and_decompresses_inert_source() {
        let source = b"Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nEnd Sub\r\n";
        let mut module_stream = vec![7, 8, 9];
        module_stream.extend_from_slice(&literal_container(source));

        let mut writer = OleWriter::new();
        writer
            .create_stream(&["PROJECT"], b"ID=\"Sample\"\r\nModule=Module1\r\n")
            .unwrap();
        writer
            .create_stream(&["VBA", "_VBA_PROJECT"], &[0; 8])
            .unwrap();
        writer
            .create_stream(&["VBA", "dir"], &sample_dir())
            .unwrap();
        writer
            .create_stream(&["VBA", "Module1"], &module_stream)
            .unwrap();
        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        cursor.set_position(0);

        let mut ole = OleFile::open(cursor).unwrap();
        let project = Project::open(&mut ole, &[], &Limits::default()).unwrap();
        assert_eq!(project.name(), "Sample");
        assert_eq!(project.page(), Mbcs::WINDOWS_1252);
        assert_eq!(project.modules().len(), 1);
        assert_eq!(project.modules()[0].source().raw(), source);
        assert_eq!(
            project.modules()[0].source().text(),
            std::str::from_utf8(source).unwrap()
        );
        assert!(!project.modules()[0].source().had_decode_errors());
        assert!(
            project
                .project_properties()
                .text()
                .contains("Module=Module1")
        );
    }

    #[test]
    fn rejects_module_offset_past_stream() {
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["PROJECT"], b"ID=\"Sample\"\r\n")
            .unwrap();
        writer
            .create_stream(&["VBA", "_VBA_PROJECT"], &[0; 8])
            .unwrap();
        writer
            .create_stream(&["VBA", "dir"], &sample_dir())
            .unwrap();
        writer.create_stream(&["VBA", "Module1"], &[1, 2]).unwrap();
        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        cursor.set_position(0);
        let mut ole = OleFile::open(cursor).unwrap();
        assert!(Project::open(&mut ole, &[], &Limits::default()).is_err());
    }

    #[test]
    fn version_stream_requires_header_but_ignores_header_values_and_cache() {
        for version_stream in [None, Some(&[1, 2, 3, 4, 5, 6][..])] {
            let mut writer = OleWriter::new();
            if let Some(stream_bytes) = version_stream {
                writer
                    .create_stream(&["VBA", "_VBA_PROJECT"], stream_bytes)
                    .unwrap();
            }
            let mut cursor = Cursor::new(Vec::new());
            writer.write_to(&mut cursor).unwrap();
            cursor.set_position(0);
            let mut ole = OleFile::open(cursor).unwrap();
            assert!(Project::open(&mut ole, &[], &Limits::default()).is_err());
        }

        let source = b"Attribute VB_Name = \"Module1\"\r\n";
        let mut module_stream = vec![7, 8, 9];
        module_stream.extend_from_slice(&literal_container(source));
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["PROJECT"], b"ID=\"Sample\"\r\nModule=Module1\r\n")
            .unwrap();
        writer
            .create_stream(
                &["VBA", "_VBA_PROJECT"],
                &[0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc, 0xde, 1, 2, 3],
            )
            .unwrap();
        writer
            .create_stream(&["VBA", "dir"], &sample_dir())
            .unwrap();
        writer
            .create_stream(&["VBA", "Module1"], &module_stream)
            .unwrap();
        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        cursor.set_position(0);
        let mut ole = OleFile::open(cursor).unwrap();
        let project = Project::open(&mut ole, &[], &Limits::default()).unwrap();
        assert_eq!(project.modules()[0].source().raw(), source);
    }
}
