//! Inert BIFF8 VBA project metadata.
//!
//! This module reports markers, object code names, and OLE storage presence.
//! It never opens, decompresses, parses, or executes VBA streams.

use super::{XlsError, XlsResult};

pub(crate) const OB_PROJ_RECORD_TYPE: u16 = 0x00D3;
pub(crate) const CODE_NAME_RECORD_TYPE: u16 = 0x01BA;
pub(crate) const OB_NO_MACROS_RECORD_TYPE: u16 = 0x01BF;
const DIMENSIONS_RECORD_TYPE: u16 = 0x0200;

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct XlsVbaMetadata {
    project_marker: bool,
    no_macros_marker: bool,
    project_storage_present: bool,
    workbook_code_name: Option<String>,
}

impl XlsVbaMetadata {
    pub fn has_project_marker(&self) -> bool {
        self.project_marker
    }
    pub fn has_no_macros_marker(&self) -> bool {
        self.no_macros_marker
    }
    pub fn has_project_storage(&self) -> bool {
        self.project_storage_present
    }
    pub fn workbook_code_name(&self) -> Option<&str> {
        self.workbook_code_name.as_deref()
    }
    pub fn may_contain_executable_code(&self) -> bool {
        self.project_marker && self.project_storage_present && !self.no_macros_marker
    }
    pub fn markers_are_consistent(&self) -> bool {
        !self.no_macros_marker || self.project_marker
    }
    pub(crate) fn set_project_storage_present(&mut self, present: bool) {
        self.project_storage_present = present;
    }
}

/// Directory-only topology for the MS-XLS `_VBA_PROJECT_CUR` storage.
///
/// MS-XLS permits at most one storage with this name and delegates its
/// contents to MS-OVBA. This model examines CFB directory names only: it
/// never opens, decompresses, parses, or executes the `PROJECT`, `dir`,
/// `_VBA_PROJECT`, SRP, or candidate module streams.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsVbaProjectStorage {
    root_storage_path: Vec<String>,
    vba_storage_path: Option<Vec<String>>,
    has_project_stream: bool,
    has_project_wm_stream: bool,
    has_project_lk_stream: bool,
    has_vba_project_stream: bool,
    has_dir_stream: bool,
    candidate_module_stream_names: Vec<String>,
    srp_stream_names: Vec<String>,
}

impl XlsVbaProjectStorage {
    /// Return the CFB path of the `_VBA_PROJECT_CUR` root storage.
    pub fn root_storage_path(&self) -> &[String] {
        &self.root_storage_path
    }

    /// Return the CFB path of the nested MS-OVBA `VBA` storage, when visible.
    ///
    /// CFB directory enumeration exposes stream paths, so an entirely empty
    /// nested storage has no observable path and is reported as absent.
    pub fn vba_storage_path(&self) -> Option<&[String]> {
        self.vba_storage_path.as_deref()
    }

    /// Whether directory metadata shows a nested `VBA` storage.
    pub fn has_vba_storage(&self) -> bool {
        self.vba_storage_path.is_some()
    }

    /// Whether the root has the required MS-OVBA `PROJECT` stream.
    pub fn has_project_stream(&self) -> bool {
        self.has_project_stream
    }

    /// Whether the root has the optional MS-OVBA `PROJECTwm` stream.
    pub fn has_project_wm_stream(&self) -> bool {
        self.has_project_wm_stream
    }

    /// Whether the root has the optional MS-OVBA `PROJECTlk` stream.
    pub fn has_project_lk_stream(&self) -> bool {
        self.has_project_lk_stream
    }

    /// Whether the nested `VBA` storage has the required `_VBA_PROJECT` stream.
    pub fn has_vba_project_stream(&self) -> bool {
        self.has_vba_project_stream
    }

    /// Whether the nested `VBA` storage has the required compressed `dir` stream.
    pub fn has_dir_stream(&self) -> bool {
        self.has_dir_stream
    }

    /// Return direct `VBA` child streams that might be module streams.
    ///
    /// `_VBA_PROJECT`, `dir`, and optional `__SRP_*` streams are excluded.
    /// Names are directory metadata only; no source bytes are read or
    /// interpreted.
    pub fn candidate_module_stream_names(&self) -> &[String] {
        &self.candidate_module_stream_names
    }

    /// Return optional `__SRP_*` streams observed in the `VBA` storage.
    ///
    /// MS-OVBA requires these streams to be ignored. This accessor returns
    /// names only and never reads their contents.
    pub fn srp_stream_names(&self) -> &[String] {
        &self.srp_stream_names
    }

    /// Whether both streams required inside the nested `VBA` storage exist.
    pub fn has_required_vba_streams(&self) -> bool {
        self.has_vba_project_stream && self.has_dir_stream
    }

    /// Whether the observed directory names meet the required MS-XLS/MS-OVBA
    /// topology.
    ///
    /// This is directory validation only. It does not validate stream bytes or
    /// parse any VBA source code.
    pub fn is_structurally_complete(&self) -> bool {
        self.has_vba_storage() && self.has_project_stream && self.has_required_vba_streams()
    }

    /// Whether directory metadata conservatively suggests candidate macro code.
    ///
    /// This is not code analysis and does not override BIFF's `ObNoMacros`
    /// marker. Candidate module stream contents are never opened,
    /// decompressed, parsed, or executed.
    pub fn may_contain_macro_code(&self) -> bool {
        self.is_structurally_complete() && !self.candidate_module_stream_names.is_empty()
    }
}

/// Discover the one MS-XLS `_VBA_PROJECT_CUR` storage from CFB directory names.
///
/// The storage name is fixed by MS-XLS, and all names defined by MS-OVBA are
/// case-insensitive. No stream content is opened by this function.
pub(crate) fn discover_vba_project_storage(
    stream_paths: &[Vec<String>],
) -> Option<XlsVbaProjectStorage> {
    let root_name = stream_paths
        .iter()
        .filter_map(|path| path.first())
        .find(|name| name.eq_ignore_ascii_case("_VBA_PROJECT_CUR"))?
        .clone();
    let root_storage_path = vec![root_name];

    let vba_storage_path = stream_paths
        .iter()
        .filter(|path| {
            path.len() >= 3
                && path_prefix_eq_ignore_ascii_case(path, &root_storage_path)
                && path[1].eq_ignore_ascii_case("VBA")
        })
        .map(|path| path[1].clone())
        .min_by(|left, right| compare_case_insensitively(left, right))
        .map(|name| vec![root_storage_path[0].clone(), name]);

    let mut vba_children = vba_storage_path
        .as_ref()
        .map(|path| direct_child_stream_names(stream_paths, path))
        .unwrap_or_default();
    sort_case_insensitively(&mut vba_children);
    vba_children.dedup_by(|left, right| left.eq_ignore_ascii_case(right));

    let has_project_stream = has_direct_stream(stream_paths, &root_storage_path, "PROJECT");
    let has_project_wm_stream = has_direct_stream(stream_paths, &root_storage_path, "PROJECTwm");
    let has_project_lk_stream = has_direct_stream(stream_paths, &root_storage_path, "PROJECTlk");
    let has_vba_project_stream = vba_children
        .iter()
        .any(|name| name.eq_ignore_ascii_case("_VBA_PROJECT"));
    let has_dir_stream = vba_children
        .iter()
        .any(|name| name.eq_ignore_ascii_case("dir"));
    let srp_stream_names = vba_children
        .iter()
        .filter(|name| is_srp_stream(name))
        .cloned()
        .collect();
    let candidate_module_stream_names = vba_children
        .into_iter()
        .filter(|name| {
            !name.eq_ignore_ascii_case("_VBA_PROJECT")
                && !name.eq_ignore_ascii_case("dir")
                && !is_srp_stream(name)
        })
        .collect();

    Some(XlsVbaProjectStorage {
        root_storage_path,
        vba_storage_path,
        has_project_stream,
        has_project_wm_stream,
        has_project_lk_stream,
        has_vba_project_stream,
        has_dir_stream,
        candidate_module_stream_names,
        srp_stream_names,
    })
}

fn has_direct_stream(stream_paths: &[Vec<String>], parent: &[String], name: &str) -> bool {
    stream_paths.iter().any(|path| {
        path.len() == parent.len() + 1
            && path_prefix_eq_ignore_ascii_case(path, parent)
            && path
                .last()
                .is_some_and(|component| component.eq_ignore_ascii_case(name))
    })
}

fn direct_child_stream_names(stream_paths: &[Vec<String>], parent: &[String]) -> Vec<String> {
    stream_paths
        .iter()
        .filter(|path| path.len() == parent.len() + 1)
        .filter(|path| path_prefix_eq_ignore_ascii_case(path, parent))
        .filter_map(|path| path.last())
        .cloned()
        .collect()
}

fn path_prefix_eq_ignore_ascii_case(path: &[String], prefix: &[String]) -> bool {
    path.iter()
        .zip(prefix)
        .all(|(component, expected)| component.eq_ignore_ascii_case(expected))
}

fn is_srp_stream(name: &str) -> bool {
    name.get(..6)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("__SRP_"))
}

fn sort_case_insensitively(names: &mut [String]) {
    names.sort_by(|left, right| compare_case_insensitively(left, right));
}

fn compare_case_insensitively(left: &str, right: &str) -> std::cmp::Ordering {
    left.to_ascii_lowercase()
        .cmp(&right.to_ascii_lowercase())
        .then_with(|| left.cmp(right))
}

pub(crate) struct WorkbookVbaCollector {
    metadata: XlsVbaMetadata,
    last_rank: Option<u8>,
}

impl WorkbookVbaCollector {
    pub(crate) fn new() -> Self {
        Self {
            metadata: XlsVbaMetadata::default(),
            last_rank: None,
        }
    }
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        let rank = match record_type {
            OB_PROJ_RECORD_TYPE => 0,
            OB_NO_MACROS_RECORD_TYPE => 1,
            CODE_NAME_RECORD_TYPE => 2,
            _ => return Ok(()),
        };
        if self.last_rank.is_some_and(|previous| rank < previous) {
            return invalid(
                record_type,
                "VBA metadata record is out of workbook-global order",
            );
        }
        self.last_rank = Some(rank);
        match record_type {
            OB_PROJ_RECORD_TYPE => {
                if self.metadata.project_marker {
                    return invalid(record_type, "duplicate ObProj record");
                }
                require_empty(record_type, data)?;
                self.metadata.project_marker = true;
            },
            OB_NO_MACROS_RECORD_TYPE => {
                if self.metadata.no_macros_marker {
                    return invalid(record_type, "duplicate ObNoMacros record");
                }
                require_empty(record_type, data)?;
                if !self.metadata.project_marker {
                    return invalid(record_type, "ObNoMacros requires a preceding ObProj record");
                }
                self.metadata.no_macros_marker = true;
            },
            CODE_NAME_RECORD_TYPE => {
                if self.metadata.workbook_code_name.is_some() {
                    return invalid(record_type, "duplicate workbook CodeName record");
                }
                self.metadata.workbook_code_name = Some(parse_code_name(data)?);
            },
            _ => unreachable!(),
        }
        Ok(())
    }
    pub(crate) fn finish(self) -> XlsVbaMetadata {
        self.metadata
    }
}

pub(crate) struct WorksheetVbaCollector {
    code_name: Option<String>,
    dimensions_seen: bool,
}

impl WorksheetVbaCollector {
    pub(crate) fn new() -> Self {
        Self {
            code_name: None,
            dimensions_seen: false,
        }
    }
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        match record_type {
            DIMENSIONS_RECORD_TYPE => self.dimensions_seen = true,
            OB_PROJ_RECORD_TYPE | OB_NO_MACROS_RECORD_TYPE => {
                return invalid(
                    record_type,
                    "workbook VBA marker appears in worksheet scope",
                );
            },
            CODE_NAME_RECORD_TYPE => {
                if !self.dimensions_seen {
                    return invalid(
                        record_type,
                        "worksheet CodeName appears before worksheet content",
                    );
                }
                if self.code_name.is_some() {
                    return invalid(record_type, "duplicate worksheet CodeName record");
                }
                self.code_name = Some(parse_code_name(data)?);
            },
            _ => {},
        }
        Ok(())
    }
    pub(crate) fn finish(self) -> Option<String> {
        self.code_name
    }
}

pub(crate) fn validate_code_name(value: &str) -> XlsResult<()> {
    if value.encode_utf16().count() > 31 {
        return invalid_data("VBA object code name exceeds 31 UTF-16 code units");
    }
    if value.is_empty() {
        return Ok(());
    }
    let mut characters = value.chars();
    let first = characters.next().unwrap();
    if first.is_ascii() && !first.is_ascii_alphabetic() {
        return invalid_data("VBA object code name must begin with a letter");
    }
    if characters.any(|character| {
        character.is_ascii() && !(character.is_ascii_alphanumeric() || character == '_')
    }) {
        return invalid_data("VBA object code name contains an invalid ASCII character");
    }
    if value.chars().any(|character| character == '\u{FFE3}') {
        return invalid_data("VBA object code name contains forbidden U+FFE3");
    }
    Ok(())
}

pub(crate) fn parse_code_name(data: &[u8]) -> XlsResult<String> {
    if data.len() < 3 {
        return invalid(CODE_NAME_RECORD_TYPE, "truncated CodeName record");
    }
    let count = u16::from_le_bytes([data[0], data[1]]) as usize;
    if count > 31 {
        return invalid(CODE_NAME_RECORD_TYPE, "CodeName exceeds 31 characters");
    }
    let options = data[2];
    if options & 0xFE != 0 {
        return invalid(
            CODE_NAME_RECORD_TYPE,
            "CodeName contains reserved string option bits",
        );
    }
    let width = if options & 1 == 0 { 1 } else { 2 };
    let expected = 3usize
        .checked_add(
            count
                .checked_mul(width)
                .ok_or_else(|| XlsError::InvalidData("CodeName size overflow".to_string()))?,
        )
        .ok_or_else(|| XlsError::InvalidData("CodeName size overflow".to_string()))?;
    if data.len() != expected {
        return invalid(
            CODE_NAME_RECORD_TYPE,
            "CodeName character count does not match payload length",
        );
    }
    let value = if width == 1 {
        data[3..].iter().map(|byte| char::from(*byte)).collect()
    } else {
        let units: Vec<u16> = data[3..]
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        String::from_utf16(&units).map_err(|_| XlsError::InvalidRecord {
            record_type: CODE_NAME_RECORD_TYPE,
            message: "CodeName contains invalid UTF-16".to_string(),
        })?
    };
    validate_code_name(&value).map_err(|error| match error {
        XlsError::InvalidData(message) => XlsError::InvalidRecord {
            record_type: CODE_NAME_RECORD_TYPE,
            message,
        },
        other => other,
    })?;
    Ok(value)
}

fn require_empty(record_type: u16, data: &[u8]) -> XlsResult<()> {
    if !data.is_empty() {
        return invalid(record_type, "marker record payload must be empty");
    }
    Ok(())
}
fn invalid<T>(record_type: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    })
}
fn invalid_data<T>(message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidData(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xls::{XlsWorkbook, XlsWriter};
    use crate::{OleFile, OleWriter};
    use std::io::Cursor;

    #[test]
    fn parses_strict_code_names() {
        assert_eq!(
            parse_code_name(&[6, 0, 0, b'S', b'h', b'e', b'e', b't', b'1']).unwrap(),
            "Sheet1"
        );
        assert!(parse_code_name(&[1, 0, 2, b'A']).is_err());
        assert!(parse_code_name(&[2, 0, 0, b'A']).is_err());
        assert!(parse_code_name(&[1, 0, 0, b'1']).is_err());
    }
    #[test]
    fn rejects_bad_marker_lengths_order_and_scope() {
        let mut globals = WorkbookVbaCollector::new();
        assert!(globals.feed_record(OB_PROJ_RECORD_TYPE, &[0]).is_err());
        let mut globals = WorkbookVbaCollector::new();
        assert!(globals.feed_record(OB_NO_MACROS_RECORD_TYPE, &[]).is_err());
        let mut sheet = WorksheetVbaCollector::new();
        assert!(sheet.feed_record(OB_PROJ_RECORD_TYPE, &[]).is_err());
        assert!(
            sheet
                .feed_record(CODE_NAME_RECORD_TYPE, &[1, 0, 0, b'A'])
                .is_err()
        );
    }
    #[test]
    fn no_macros_marker_is_inert() {
        let mut globals = WorkbookVbaCollector::new();
        globals.feed_record(OB_PROJ_RECORD_TYPE, &[]).unwrap();
        globals.feed_record(OB_NO_MACROS_RECORD_TYPE, &[]).unwrap();
        let metadata = globals.finish();
        assert!(metadata.has_project_marker());
        assert!(metadata.has_no_macros_marker());
        assert!(!metadata.may_contain_executable_code());
    }

    #[test]
    fn discovers_xls_vba_storage_topology_from_directory_names_only() {
        let stream_paths = vec![
            vec!["Workbook".to_string()],
            vec!["_VBA_PROJECT_CUR".to_string(), "PROJECT".to_string()],
            vec!["_VBA_PROJECT_CUR".to_string(), "PROJECTwm".to_string()],
            vec!["_VBA_PROJECT_CUR".to_string(), "PROJECTlk".to_string()],
            vec![
                "_VBA_PROJECT_CUR".to_string(),
                "vBa".to_string(),
                "_vba_project".to_string(),
            ],
            vec![
                "_VBA_PROJECT_CUR".to_string(),
                "vBa".to_string(),
                "DIR".to_string(),
            ],
            vec![
                "_VBA_PROJECT_CUR".to_string(),
                "vBa".to_string(),
                "ThisWorkbook".to_string(),
            ],
            vec![
                "_VBA_PROJECT_CUR".to_string(),
                "vBa".to_string(),
                "Module1".to_string(),
            ],
            vec![
                "_VBA_PROJECT_CUR".to_string(),
                "vBa".to_string(),
                "__sRp_0".to_string(),
            ],
        ];

        let storage = discover_vba_project_storage(&stream_paths).unwrap();
        assert_eq!(storage.root_storage_path(), ["_VBA_PROJECT_CUR"]);
        assert_eq!(
            storage.vba_storage_path().unwrap(),
            ["_VBA_PROJECT_CUR", "vBa"]
        );
        assert!(storage.has_project_stream());
        assert!(storage.has_project_wm_stream());
        assert!(storage.has_project_lk_stream());
        assert!(storage.has_required_vba_streams());
        assert!(storage.is_structurally_complete());
        assert!(storage.may_contain_macro_code());
        assert_eq!(
            storage.candidate_module_stream_names(),
            ["Module1", "ThisWorkbook"]
        );
        assert_eq!(storage.srp_stream_names(), ["__sRp_0"]);
    }

    #[test]
    fn parsed_xls_workbook_discovers_invalid_vba_payloads_as_inert_metadata() {
        let mut source_writer = XlsWriter::new();
        source_writer.add_worksheet("Sheet1").unwrap();
        let mut source_bytes = Cursor::new(Vec::new());
        source_writer.write_to(&mut source_bytes).unwrap();
        let mut source = OleFile::open(Cursor::new(source_bytes.into_inner())).unwrap();
        let workbook_stream = source.open_stream(&["Workbook"]).unwrap();

        let mut writer = OleWriter::new();
        writer
            .create_stream(&["Workbook"], &workbook_stream)
            .unwrap();
        writer.create_storage(&["_VBA_PROJECT_CUR"]).unwrap();
        writer
            .create_stream(
                &["_VBA_PROJECT_CUR", "PROJECT"],
                b"intentionally invalid PROJECT bytes",
            )
            .unwrap();
        writer.create_storage(&["_VBA_PROJECT_CUR", "VBA"]).unwrap();
        writer
            .create_stream(
                &["_VBA_PROJECT_CUR", "VBA", "_VBA_PROJECT"],
                b"intentionally invalid version data",
            )
            .unwrap();
        writer
            .create_stream(
                &["_VBA_PROJECT_CUR", "VBA", "dir"],
                b"intentionally not an OVBA compressed container",
            )
            .unwrap();
        writer
            .create_stream(
                &["_VBA_PROJECT_CUR", "VBA", "Module1"],
                b"intentionally not a module stream",
            )
            .unwrap();
        writer
            .create_stream(
                &["_VBA_PROJECT_CUR", "VBA", "__SRP_0"],
                b"intentionally ignored",
            )
            .unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();

        let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
        let storage = workbook.vba_project_storage().unwrap();
        assert!(storage.is_structurally_complete());
        assert!(storage.may_contain_macro_code());
        assert_eq!(storage.candidate_module_stream_names(), ["Module1"]);
        assert_eq!(storage.srp_stream_names(), ["__SRP_0"]);
        assert!(workbook.vba_metadata().has_project_storage());
        assert!(!workbook.vba_metadata().has_project_marker());
    }

    #[test]
    fn parses_real_xls_vba_project_and_inert_module_source() {
        let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet/SimpleMacro.xls");
        let mut workbook = XlsWorkbook::new(std::fs::File::open(fixture).unwrap()).unwrap();
        let project = workbook
            .vba_project(&litchi_cfb::ovba::VbaLimits::default())
            .unwrap()
            .unwrap();

        assert!(!project.name().is_empty());
        assert!(!project.modules().is_empty());
        assert!(
            project
                .modules()
                .iter()
                .any(|module| module.source().text().contains("Sub "))
        );
    }
}
