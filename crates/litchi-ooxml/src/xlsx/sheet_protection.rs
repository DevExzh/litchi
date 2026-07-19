//! Immutable worksheet-protection metadata for SpreadsheetML worksheets.

use std::collections::HashSet;
use std::fmt::Write;
use std::ops::Range;

use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::common::{ExpandedName, MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};

const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM: &[u8] = b"http://schemas.microsoft.com/office/excel/2006/main";
const PROTECTED_RANGES_EXTENSION_URI: &str = "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_RANGES: usize = 100_000;
const MAX_REFERENCES: usize = 8_192;
const MAX_STRING_BYTES: usize = 32 * 1024;
const MAX_BINARY_BYTES: usize = 1024 * 1024;
const MAX_SPIN_COUNT: u32 = 10_000_000;
const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;

/// SpreadsheetML namespace form used by the deterministic writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WorksheetProtectionConformance {
    Transitional,
    Strict,
}

impl WorksheetProtectionConformance {
    fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => std::str::from_utf8(CORE).unwrap(),
            Self::Strict => std::str::from_utf8(STRICT).unwrap(),
        }
    }
}

/// Source schema for a protected-range collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ProtectedRangeSource {
    /// ISO/IEC 29500 SpreadsheetML collection.
    Core,
    /// Office 2010 `x14` worksheet extension.
    Office2010,
}

/// Password-verifier metadata. This type does not verify passwords.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ProtectionPasswordVerifier {
    /// Legacy 16-bit SpreadsheetML password verifier.
    Legacy(u16),
    /// Salted iterative password hash metadata.
    Strong(StrongProtectionPasswordVerifier),
}

/// Salted iterative password hash metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StrongProtectionPasswordVerifier {
    algorithm_name: String,
    hash_value: Vec<u8>,
    salt_value: Vec<u8>,
    spin_count: u32,
}

impl StrongProtectionPasswordVerifier {
    pub fn new(
        algorithm_name: impl Into<String>,
        hash_value: Vec<u8>,
        salt_value: Vec<u8>,
        spin_count: u32,
    ) -> Result<Self> {
        let value = Self {
            algorithm_name: algorithm_name.into(),
            hash_value,
            salt_value,
            spin_count,
        };
        validate_strong_verifier(&value)?;
        Ok(value)
    }

    pub fn algorithm_name(&self) -> &str {
        &self.algorithm_name
    }
    pub fn hash_value(&self) -> &[u8] {
        &self.hash_value
    }
    pub fn salt_value(&self) -> &[u8] {
        &self.salt_value
    }
    pub fn spin_count(&self) -> u32 {
        self.spin_count
    }
}

/// Effective operation locks from a worksheet `sheetProtection` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetProtection {
    verifier: Option<ProtectionPasswordVerifier>,
    sheet: bool,
    objects: bool,
    scenarios: bool,
    format_cells: bool,
    format_columns: bool,
    format_rows: bool,
    insert_columns: bool,
    insert_rows: bool,
    insert_hyperlinks: bool,
    delete_columns: bool,
    delete_rows: bool,
    select_locked_cells: bool,
    sort: bool,
    auto_filter: bool,
    pivot_tables: bool,
    select_unlocked_cells: bool,
}

impl WorksheetProtection {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn verifier(&self) -> Option<&ProtectionPasswordVerifier> {
        self.verifier.as_ref()
    }
    pub fn sheet_locked(&self) -> bool {
        self.sheet
    }
    pub fn objects_locked(&self) -> bool {
        self.objects
    }
    pub fn scenarios_locked(&self) -> bool {
        self.scenarios
    }
    pub fn format_cells_locked(&self) -> bool {
        self.format_cells
    }
    pub fn format_columns_locked(&self) -> bool {
        self.format_columns
    }
    pub fn format_rows_locked(&self) -> bool {
        self.format_rows
    }
    pub fn insert_columns_locked(&self) -> bool {
        self.insert_columns
    }
    pub fn insert_rows_locked(&self) -> bool {
        self.insert_rows
    }
    pub fn insert_hyperlinks_locked(&self) -> bool {
        self.insert_hyperlinks
    }
    pub fn delete_columns_locked(&self) -> bool {
        self.delete_columns
    }
    pub fn delete_rows_locked(&self) -> bool {
        self.delete_rows
    }
    pub fn select_locked_cells_locked(&self) -> bool {
        self.select_locked_cells
    }
    pub fn sort_locked(&self) -> bool {
        self.sort
    }
    pub fn auto_filter_locked(&self) -> bool {
        self.auto_filter
    }
    pub fn pivot_tables_locked(&self) -> bool {
        self.pivot_tables
    }
    pub fn select_unlocked_cells_locked(&self) -> bool {
        self.select_unlocked_cells
    }

    pub fn set_verifier(&mut self, verifier: Option<ProtectionPasswordVerifier>) -> Result<()> {
        if let Some(ProtectionPasswordVerifier::Strong(value)) = verifier.as_ref() {
            validate_strong_verifier(value)?;
        }
        self.verifier = verifier;
        Ok(())
    }

    pub fn set_sheet_locked(&mut self, value: bool) { self.sheet = value; }
    pub fn set_objects_locked(&mut self, value: bool) { self.objects = value; }
    pub fn set_scenarios_locked(&mut self, value: bool) { self.scenarios = value; }
    pub fn set_format_cells_locked(&mut self, value: bool) { self.format_cells = value; }
    pub fn set_format_columns_locked(&mut self, value: bool) { self.format_columns = value; }
    pub fn set_format_rows_locked(&mut self, value: bool) { self.format_rows = value; }
    pub fn set_insert_columns_locked(&mut self, value: bool) { self.insert_columns = value; }
    pub fn set_insert_rows_locked(&mut self, value: bool) { self.insert_rows = value; }
    pub fn set_insert_hyperlinks_locked(&mut self, value: bool) { self.insert_hyperlinks = value; }
    pub fn set_delete_columns_locked(&mut self, value: bool) { self.delete_columns = value; }
    pub fn set_delete_rows_locked(&mut self, value: bool) { self.delete_rows = value; }
    pub fn set_select_locked_cells_locked(&mut self, value: bool) { self.select_locked_cells = value; }
    pub fn set_sort_locked(&mut self, value: bool) { self.sort = value; }
    pub fn set_auto_filter_locked(&mut self, value: bool) { self.auto_filter = value; }
    pub fn set_pivot_tables_locked(&mut self, value: bool) { self.pivot_tables = value; }
    pub fn set_select_unlocked_cells_locked(&mut self, value: bool) { self.select_unlocked_cells = value; }
}

impl Default for WorksheetProtection {
    fn default() -> Self {
        Self {
            verifier: None,
            sheet: false,
            objects: false,
            scenarios: false,
            format_cells: true,
            format_columns: true,
            format_rows: true,
            insert_columns: true,
            insert_rows: true,
            insert_hyperlinks: true,
            delete_columns: true,
            delete_rows: true,
            select_locked_cells: false,
            sort: true,
            auto_filter: true,
            pivot_tables: true,
            select_unlocked_cells: false,
        }
    }
}

/// Typed kind of an individual protected-range reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionRangeReferenceKind {
    Cells {
        start_row: u32,
        start_column: u32,
        end_row: u32,
        end_column: u32,
    },
    Columns {
        start_column: u32,
        end_column: u32,
    },
    Rows {
        start_row: u32,
        end_row: u32,
    },
}

/// One validated reference in a protected range's `sqref`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionRangeReference {
    raw: String,
    kind: ProtectionRangeReferenceKind,
}

impl ProtectionRangeReference {
    pub fn raw(&self) -> &str {
        &self.raw
    }
    pub fn kind(&self) -> ProtectionRangeReferenceKind {
        self.kind
    }
}

/// Validated whitespace-separated protected-range references.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionRangeSqref {
    raw: String,
    references: Vec<ProtectionRangeReference>,
}

impl ProtectionRangeSqref {
    pub fn parse(value: impl AsRef<str>) -> Result<Self> {
        parse_sqref(value.as_ref())
    }

    pub fn raw(&self) -> &str {
        &self.raw
    }
    pub fn references(&self) -> &[ProtectionRangeReference] {
        &self.references
    }
}

/// A single editable range associated with worksheet protection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetProtectedRange {
    source: ProtectedRangeSource,
    name: String,
    sqref: ProtectionRangeSqref,
    verifier: Option<ProtectionPasswordVerifier>,
    security_descriptor: Option<String>,
}

impl WorksheetProtectedRange {
    pub fn new(
        source: ProtectedRangeSource,
        name: impl Into<String>,
        sqref: ProtectionRangeSqref,
    ) -> Result<Self> {
        let value = Self {
            source,
            name: name.into(),
            sqref,
            verifier: None,
            security_descriptor: None,
        };
        validate_range(&value)?;
        Ok(value)
    }

    pub fn set_verifier(&mut self, verifier: Option<ProtectionPasswordVerifier>) -> Result<()> {
        if let Some(ProtectionPasswordVerifier::Strong(value)) = verifier.as_ref() {
            validate_strong_verifier(value)?;
        }
        self.verifier = verifier;
        Ok(())
    }

    pub fn set_security_descriptor(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = value.as_deref() {
            bounded(value, "securityDescriptor")?;
            validate_xml_text(value, "securityDescriptor")?;
        }
        self.security_descriptor = value;
        Ok(())
    }

    pub fn source(&self) -> ProtectedRangeSource {
        self.source
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn sqref(&self) -> &ProtectionRangeSqref {
        &self.sqref
    }
    pub fn verifier(&self) -> Option<&ProtectionPasswordVerifier> {
        self.verifier.as_ref()
    }
    pub fn security_descriptor(&self) -> Option<&str> {
        self.security_descriptor.as_deref()
    }
}

/// A protected-range container in worksheet document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetProtectedRangeCollection {
    source: ProtectedRangeSource,
    ranges: Vec<WorksheetProtectedRange>,
}

impl WorksheetProtectedRangeCollection {
    pub fn new(
        source: ProtectedRangeSource,
        ranges: Vec<WorksheetProtectedRange>,
    ) -> Result<Self> {
        let value = Self { source, ranges };
        validate_collection(&value)?;
        Ok(value)
    }

    pub fn source(&self) -> ProtectedRangeSource {
        self.source
    }
    pub fn ranges(&self) -> &[WorksheetProtectedRange] {
        &self.ranges
    }
}

/// Complete worksheet protection metadata.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorksheetProtectionMetadata {
    sheet_protection: Option<WorksheetProtection>,
    protected_range_collections: Vec<WorksheetProtectedRangeCollection>,
}

impl WorksheetProtectionMetadata {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn sheet_protection(&self) -> Option<&WorksheetProtection> {
        self.sheet_protection.as_ref()
    }
    pub fn protected_range_collections(&self) -> &[WorksheetProtectedRangeCollection] {
        &self.protected_range_collections
    }
    pub fn protected_ranges(&self) -> impl Iterator<Item = &WorksheetProtectedRange> {
        self.protected_range_collections
            .iter()
            .flat_map(|collection| collection.ranges.iter())
    }

    pub fn set_sheet_protection(&mut self, value: Option<WorksheetProtection>) -> Result<()> {
        if let Some(value) = value.as_ref() {
            validate_sheet_protection(value)?;
        }
        self.sheet_protection = value;
        Ok(())
    }

    pub fn set_protected_range_collections(
        &mut self,
        value: Vec<WorksheetProtectedRangeCollection>,
    ) -> Result<()> {
        let candidate = Self {
            sheet_protection: self.sheet_protection.clone(),
            protected_range_collections: value,
        };
        validate_worksheet_protection_metadata(&candidate)?;
        self.protected_range_collections = candidate.protected_range_collections;
        Ok(())
    }

    pub fn clear_sheet_protection(&mut self) {
        self.sheet_protection = None;
    }

    pub fn clear_protected_ranges(&mut self) {
        self.protected_range_collections.clear();
    }
}

#[derive(Default)]
struct RawCredential {
    password: Option<String>,
    algorithm_name: Option<String>,
    hash_value: Option<String>,
    salt_value: Option<String>,
    spin_count: Option<String>,
}

struct PendingRange {
    source: ProtectedRangeSource,
    name: String,
    sqref: Option<String>,
    credential: RawCredential,
    security_descriptor: Option<String>,
}

/// Parse protection metadata from a complete worksheet XML part.
pub fn parse_worksheet_protection(xml: &[u8]) -> Result<WorksheetProtectionMetadata> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML is too large"));
    }
    let mut capabilities = MceCapabilities::default();
    capabilities
        .understand_namespace(String::from_utf8_lossy(X14).into_owned())
        .understand_namespace(String::from_utf8_lossy(XM).into_owned());
    capabilities.preserve_extension_element(ExpandedName {
        namespace: String::from_utf8_lossy(X14).into_owned(),
        local_name: "protectedRanges".into(),
    });
    let validated = process_markup_compatibility(xml, &capabilities, &MceLimits::default())?;
    parse_selected(validated.xml.as_ref())
}

fn parse_selected(xml: &[u8]) -> Result<WorksheetProtectionMetadata> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut metadata = WorksheetProtectionMetadata::default();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut extension_depth = None;
    let mut core_collection: Option<(usize, Vec<WorksheetProtectedRange>)> = None;
    let mut x14_collection: Option<(usize, Vec<WorksheetProtectedRange>)> = None;
    let mut pending: Option<(usize, PendingRange)> = None;
    let mut sqref_text: Option<(usize, String)> = None;
    let mut sheet_protection_depth = None;
    let mut worksheet_order = 0u8;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                if depth == 1 {
                    if root_seen
                        || !spreadsheet(&namespace)
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid(
                            "worksheet protection parser requires a worksheet root",
                        ));
                    }
                    root_seen = true;
                    continue;
                }
                if depth == 2 {
                    update_worksheet_order(
                        &namespace,
                        element.local_name().as_ref(),
                        &mut worksheet_order,
                    )?;
                }
                if sheet_protection_depth.is_some() {
                    return Err(invalid("sheetProtection must be empty"));
                }
                if let Some((sqref_depth, _)) = sqref_text.as_ref() {
                    if depth > *sqref_depth {
                        return Err(invalid("protected-range sqref must contain only text"));
                    }
                }
                if spreadsheet(&namespace) && element.local_name().as_ref() == b"ext" {
                    if attribute(&element, decoder, &resolver, b"uri")?.as_deref()
                        == Some(PROTECTED_RANGES_EXTENSION_URI)
                    {
                        extension_depth = Some(depth);
                    }
                } else if depth == 2
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"sheetProtection"
                {
                    if metadata.sheet_protection.is_some() {
                        return Err(invalid("duplicate sheetProtection element"));
                    }
                    metadata.sheet_protection =
                        Some(parse_sheet_protection(&element, decoder, &resolver)?);
                    sheet_protection_depth = Some(depth);
                } else if depth == 2
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"protectedRanges"
                {
                    if core_collection.is_some()
                        || metadata
                            .protected_range_collections
                            .iter()
                            .any(|c| c.source == ProtectedRangeSource::Core)
                    {
                        return Err(invalid("duplicate core protectedRanges element"));
                    }
                    core_collection = Some((depth, Vec::new()));
                } else if exact(&namespace, X14)
                    && element.local_name().as_ref() == b"protectedRanges"
                    && extension_depth.is_some()
                {
                    if x14_collection.is_some()
                        || metadata
                            .protected_range_collections
                            .iter()
                            .any(|c| c.source == ProtectedRangeSource::Office2010)
                    {
                        return Err(invalid("duplicate x14 protectedRanges element"));
                    }
                    x14_collection = Some((depth, Vec::new()));
                } else if element.local_name().as_ref() == b"protectedRange" {
                    let source =
                        collection_source(depth, &namespace, &core_collection, &x14_collection)?;
                    if pending.is_some() {
                        return Err(invalid("nested protectedRange element"));
                    }
                    pending = Some((
                        depth,
                        parse_pending_range(&element, decoder, &resolver, source)?,
                    ));
                } else if exact(&namespace, XM)
                    && element.local_name().as_ref() == b"sqref"
                    && x14_collection.is_some()
                {
                    let Some((range_depth, range)) = pending.as_ref() else {
                        return Err(invalid("x14 sqref is outside protectedRange"));
                    };
                    if range.source != ProtectedRangeSource::Office2010
                        || depth != *range_depth + 1
                        || sqref_text.is_some()
                    {
                        return Err(invalid("invalid x14 protected-range sqref placement"));
                    }
                    sqref_text = Some((depth, String::new()));
                } else if pending.is_some() {
                    return Err(invalid("unexpected child in protectedRange"));
                } else if direct_collection_child(depth, &core_collection, &x14_collection) {
                    return Err(invalid("unknown protectedRanges child element"));
                }
            },
            Event::Empty(element) => {
                if sheet_protection_depth.is_some() {
                    return Err(invalid("sheetProtection must be empty"));
                }
                if depth == 1 {
                    update_worksheet_order(
                        &namespace,
                        element.local_name().as_ref(),
                        &mut worksheet_order,
                    )?;
                }
                if depth == 1
                    && spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"sheetProtection"
                {
                    if metadata.sheet_protection.is_some() {
                        return Err(invalid("duplicate sheetProtection element"));
                    }
                    metadata.sheet_protection =
                        Some(parse_sheet_protection(&element, decoder, &resolver)?);
                } else if element.local_name().as_ref() == b"protectedRange" {
                    let source = collection_source(
                        depth + 1,
                        &namespace,
                        &core_collection,
                        &x14_collection,
                    )?;
                    let range =
                        finish_range(parse_pending_range(&element, decoder, &resolver, source)?)?;
                    push_range(range, &mut core_collection, &mut x14_collection)?;
                } else if exact(&namespace, XM)
                    && element.local_name().as_ref() == b"sqref"
                    && x14_collection.is_some()
                {
                    if pending.is_some() {
                        return Err(invalid("x14 protected-range sqref cannot be empty"));
                    }
                    return Err(invalid("x14 sqref is outside protectedRange"));
                } else if direct_collection_child(depth + 1, &core_collection, &x14_collection) {
                    return Err(invalid("unknown protectedRanges child element"));
                }
            },
            Event::Text(text) => {
                if let Some((_, value)) = sqref_text.as_mut() {
                    let decoded = text.decode().map_err(xml_error)?;
                    if value.len().saturating_add(decoded.len()) > MAX_STRING_BYTES {
                        return Err(invalid("protected-range sqref is too large"));
                    }
                    value.push_str(&decoded);
                } else if (sheet_protection_depth.is_some()
                    || pending.is_some()
                    || core_collection.is_some()
                    || x14_collection.is_some())
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("protection elements cannot contain text"));
                }
            },
            Event::End(element) => {
                if sqref_text
                    .as_ref()
                    .is_some_and(|(sqref_depth, _)| *sqref_depth == depth)
                {
                    let (_, value) = sqref_text.take().expect("checked above");
                    let Some((_, range)) = pending.as_mut() else {
                        return Err(invalid("orphan x14 sqref"));
                    };
                    if range.sqref.replace(value).is_some() {
                        return Err(invalid("duplicate x14 protected-range sqref"));
                    }
                } else if pending
                    .as_ref()
                    .is_some_and(|(range_depth, _)| *range_depth == depth)
                {
                    let (_, range) = pending.take().expect("checked above");
                    push_range(
                        finish_range(range)?,
                        &mut core_collection,
                        &mut x14_collection,
                    )?;
                } else if core_collection
                    .as_ref()
                    .is_some_and(|(collection_depth, _)| *collection_depth == depth)
                {
                    let (_, ranges) = core_collection.take().expect("checked above");
                    if ranges.is_empty() {
                        return Err(invalid(
                            "protectedRanges must contain at least one protectedRange",
                        ));
                    }
                    metadata
                        .protected_range_collections
                        .push(WorksheetProtectedRangeCollection {
                            source: ProtectedRangeSource::Core,
                            ranges,
                        });
                } else if x14_collection
                    .as_ref()
                    .is_some_and(|(collection_depth, _)| *collection_depth == depth)
                {
                    let (_, ranges) = x14_collection.take().expect("checked above");
                    if ranges.is_empty() {
                        return Err(invalid(
                            "x14 protectedRanges must contain at least one protectedRange",
                        ));
                    }
                    metadata
                        .protected_range_collections
                        .push(WorksheetProtectedRangeCollection {
                            source: ProtectedRangeSource::Office2010,
                            ranges,
                        });
                }
                if sheet_protection_depth == Some(depth) {
                    sheet_protection_depth = None;
                }
                if extension_depth == Some(depth) {
                    extension_depth = None;
                }
                if depth == 1 {
                    if !spreadsheet(&namespace) || element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("invalid worksheet closing element"));
                    }
                    root_closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unexpected XML end element"))?;
            },
            Event::CData(_) => {
                return Err(invalid(
                    "CDATA is not allowed in worksheet protection metadata",
                ));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") {
                    return Err(invalid("custom XML entities are rejected"));
                }
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !root_seen
        || !root_closed
        || depth != 0
        || pending.is_some()
        || core_collection.is_some()
        || x14_collection.is_some()
    {
        return Err(invalid("incomplete worksheet protection XML"));
    }
    Ok(metadata)
}

fn collection_source(
    element_depth: usize,
    namespace: &ResolveResult<'_>,
    core: &Option<(usize, Vec<WorksheetProtectedRange>)>,
    x14: &Option<(usize, Vec<WorksheetProtectedRange>)>,
) -> Result<ProtectedRangeSource> {
    if core
        .as_ref()
        .is_some_and(|(depth, _)| element_depth == *depth + 1)
        && spreadsheet(namespace)
    {
        Ok(ProtectedRangeSource::Core)
    } else if x14
        .as_ref()
        .is_some_and(|(depth, _)| element_depth == *depth + 1)
        && exact(namespace, X14)
    {
        Ok(ProtectedRangeSource::Office2010)
    } else {
        Err(invalid("protectedRange is outside a matching collection"))
    }
}

fn push_range(
    range: WorksheetProtectedRange,
    core: &mut Option<(usize, Vec<WorksheetProtectedRange>)>,
    x14: &mut Option<(usize, Vec<WorksheetProtectedRange>)>,
) -> Result<()> {
    let ranges = match range.source {
        ProtectedRangeSource::Core => {
            &mut core
                .as_mut()
                .ok_or_else(|| invalid("missing core protectedRanges parent"))?
                .1
        },
        ProtectedRangeSource::Office2010 => {
            &mut x14
                .as_mut()
                .ok_or_else(|| invalid("missing x14 protectedRanges parent"))?
                .1
        },
    };
    if ranges.len() >= MAX_RANGES {
        return Err(invalid("too many protected ranges"));
    }
    ranges.push(range);
    Ok(())
}

fn parse_sheet_protection(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<WorksheetProtection> {
    let mut value = WorksheetProtection::default();
    let mut credential = RawCredential::default();
    let mut seen = HashSet::new();
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        if namespace_declaration(attr.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attr.key);
        require_unqualified_attribute(&namespace, "sheetProtection")?;
        if !seen.insert(local.as_ref().to_vec()) {
            return Err(invalid("duplicate sheetProtection attribute"));
        }
        let text = attr
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        match local.as_ref() {
            b"password" => set_once(&mut credential.password, text, "password")?,
            b"algorithmName" => set_once(&mut credential.algorithm_name, text, "algorithmName")?,
            b"hashValue" => set_once(&mut credential.hash_value, text, "hashValue")?,
            b"saltValue" => set_once(&mut credential.salt_value, text, "saltValue")?,
            b"spinCount" => set_once(&mut credential.spin_count, text, "spinCount")?,
            b"sheet" => value.sheet = parse_bool(&text, "sheet")?,
            b"objects" => value.objects = parse_bool(&text, "objects")?,
            b"scenarios" => value.scenarios = parse_bool(&text, "scenarios")?,
            b"formatCells" => value.format_cells = parse_bool(&text, "formatCells")?,
            b"formatColumns" => value.format_columns = parse_bool(&text, "formatColumns")?,
            b"formatRows" => value.format_rows = parse_bool(&text, "formatRows")?,
            b"insertColumns" => value.insert_columns = parse_bool(&text, "insertColumns")?,
            b"insertRows" => value.insert_rows = parse_bool(&text, "insertRows")?,
            b"insertHyperlinks" => value.insert_hyperlinks = parse_bool(&text, "insertHyperlinks")?,
            b"deleteColumns" => value.delete_columns = parse_bool(&text, "deleteColumns")?,
            b"deleteRows" => value.delete_rows = parse_bool(&text, "deleteRows")?,
            b"selectLockedCells" => {
                value.select_locked_cells = parse_bool(&text, "selectLockedCells")?
            },
            b"sort" => value.sort = parse_bool(&text, "sort")?,
            b"autoFilter" => value.auto_filter = parse_bool(&text, "autoFilter")?,
            b"pivotTables" => value.pivot_tables = parse_bool(&text, "pivotTables")?,
            b"selectUnlockedCells" => {
                value.select_unlocked_cells = parse_bool(&text, "selectUnlockedCells")?
            },
            other => {
                return Err(invalid(format!(
                    "unknown sheetProtection attribute '{}'",
                    String::from_utf8_lossy(other)
                )));
            },
        }
    }
    value.verifier = finish_credential(credential)?;
    Ok(value)
}

fn parse_pending_range(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    source: ProtectedRangeSource,
) -> Result<PendingRange> {
    let mut name = None;
    let mut sqref = None;
    let mut security_descriptor = None;
    let mut credential = RawCredential::default();
    let mut seen = HashSet::new();
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        if namespace_declaration(attr.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attr.key);
        require_unqualified_attribute(&namespace, "protectedRange")?;
        if !seen.insert(local.as_ref().to_vec()) {
            return Err(invalid("duplicate protectedRange attribute"));
        }
        let text = attr
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        match local.as_ref() {
            b"name" => set_once(&mut name, text, "name")?,
            b"sqref" if source == ProtectedRangeSource::Core => {
                set_once(&mut sqref, text, "sqref")?
            },
            b"password" => set_once(&mut credential.password, text, "password")?,
            b"algorithmName" => set_once(&mut credential.algorithm_name, text, "algorithmName")?,
            b"hashValue" => set_once(&mut credential.hash_value, text, "hashValue")?,
            b"saltValue" => set_once(&mut credential.salt_value, text, "saltValue")?,
            b"spinCount" => set_once(&mut credential.spin_count, text, "spinCount")?,
            b"securityDescriptor" => {
                set_once(&mut security_descriptor, text, "securityDescriptor")?
            },
            other => {
                return Err(invalid(format!(
                    "unknown protectedRange attribute '{}'",
                    String::from_utf8_lossy(other)
                )));
            },
        }
    }
    let name = name.ok_or_else(|| invalid("protectedRange is missing name"))?;
    bounded_nonempty(&name, "protectedRange name")?;
    if let Some(value) = security_descriptor.as_ref() {
        bounded(value, "securityDescriptor")?;
    }
    Ok(PendingRange {
        source,
        name,
        sqref,
        credential,
        security_descriptor,
    })
}

fn finish_range(range: PendingRange) -> Result<WorksheetProtectedRange> {
    let sqref = range
        .sqref
        .ok_or_else(|| invalid("protectedRange is missing sqref"))?;
    Ok(WorksheetProtectedRange {
        source: range.source,
        name: range.name,
        sqref: parse_sqref(&sqref)?,
        verifier: finish_credential(range.credential)?,
        security_descriptor: range.security_descriptor,
    })
}

fn finish_credential(raw: RawCredential) -> Result<Option<ProtectionPasswordVerifier>> {
    let strong_present = raw.algorithm_name.is_some()
        || raw.hash_value.is_some()
        || raw.salt_value.is_some()
        || raw.spin_count.is_some();
    if raw.password.is_some() && strong_present {
        return Err(invalid(
            "legacy password and strong hash metadata are mutually exclusive",
        ));
    }
    if let Some(password) = raw.password {
        if password.len() != 4 || !password.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(invalid(
                "legacy password verifier must be four hexadecimal digits",
            ));
        }
        return Ok(Some(ProtectionPasswordVerifier::Legacy(
            u16::from_str_radix(&password, 16)
                .map_err(|_| invalid("invalid legacy password verifier"))?,
        )));
    }
    if !strong_present {
        return Ok(None);
    }
    let algorithm_name = raw
        .algorithm_name
        .ok_or_else(|| invalid("strong verifier is missing algorithmName"))?;
    bounded_nonempty(&algorithm_name, "algorithmName")?;
    let hash_value = decode_base64(
        &raw.hash_value
            .ok_or_else(|| invalid("strong verifier is missing hashValue"))?,
        "hashValue",
    )?;
    let salt_value = decode_base64(
        &raw.salt_value
            .ok_or_else(|| invalid("strong verifier is missing saltValue"))?,
        "saltValue",
    )?;
    let spin_count = raw
        .spin_count
        .ok_or_else(|| invalid("strong verifier is missing spinCount"))?
        .parse::<u32>()
        .map_err(|_| invalid("invalid spinCount"))?;
    if spin_count > MAX_SPIN_COUNT {
        return Err(invalid("spinCount exceeds 10000000"));
    }
    Ok(Some(ProtectionPasswordVerifier::Strong(
        StrongProtectionPasswordVerifier {
            algorithm_name,
            hash_value,
            salt_value,
            spin_count,
        },
    )))
}

fn parse_sqref(value: &str) -> Result<ProtectionRangeSqref> {
    bounded_nonempty(value, "protectedRange sqref")?;
    let tokens: Vec<_> = value.split_ascii_whitespace().collect();
    if tokens.is_empty() || tokens.len() > MAX_REFERENCES {
        return Err(invalid("invalid protectedRange sqref reference count"));
    }
    let references = tokens
        .into_iter()
        .map(parse_reference)
        .collect::<Result<Vec<_>>>()?;
    Ok(ProtectionRangeSqref {
        raw: value.to_string(),
        references,
    })
}

fn parse_reference(raw: &str) -> Result<ProtectionRangeReference> {
    let parts: Vec<_> = raw.split(':').collect();
    let kind = match parts.as_slice() {
        [single] => match parse_endpoint(single)? {
            Endpoint::Cell(row, column) => ProtectionRangeReferenceKind::Cells {
                start_row: row,
                start_column: column,
                end_row: row,
                end_column: column,
            },
            _ => return Err(invalid("single protected-range reference must be a cell")),
        },
        [start, end] => match (parse_endpoint(start)?, parse_endpoint(end)?) {
            (Endpoint::Cell(sr, sc), Endpoint::Cell(er, ec)) if sr <= er && sc <= ec => {
                ProtectionRangeReferenceKind::Cells {
                    start_row: sr,
                    start_column: sc,
                    end_row: er,
                    end_column: ec,
                }
            },
            (Endpoint::Column(sc), Endpoint::Column(ec)) if sc <= ec => {
                ProtectionRangeReferenceKind::Columns {
                    start_column: sc,
                    end_column: ec,
                }
            },
            (Endpoint::Row(sr), Endpoint::Row(er)) if sr <= er => {
                ProtectionRangeReferenceKind::Rows {
                    start_row: sr,
                    end_row: er,
                }
            },
            _ => return Err(invalid("invalid or reversed protected-range reference")),
        },
        _ => {
            return Err(invalid(
                "protected-range reference contains too many colons",
            ));
        },
    };
    Ok(ProtectionRangeReference {
        raw: raw.to_string(),
        kind,
    })
}

enum Endpoint {
    Cell(u32, u32),
    Column(u32),
    Row(u32),
}

fn parse_endpoint(value: &str) -> Result<Endpoint> {
    let value = value.strip_prefix('$').unwrap_or(value);
    if value.is_empty() {
        return Err(invalid("empty protected-range endpoint"));
    }
    let letter_count = value.bytes().take_while(u8::is_ascii_alphabetic).count();
    if letter_count > 0 {
        let column_text = &value[..letter_count];
        let rest = value[letter_count..]
            .strip_prefix('$')
            .unwrap_or(&value[letter_count..]);
        let column = parse_column(column_text)?;
        if rest.is_empty() {
            return Ok(Endpoint::Column(column));
        }
        let row = parse_row(rest)?;
        Ok(Endpoint::Cell(row, column))
    } else {
        Ok(Endpoint::Row(parse_row(value)?))
    }
}

fn parse_column(value: &str) -> Result<u32> {
    let mut column = 0u32;
    for byte in value.bytes() {
        if !byte.is_ascii_alphabetic() {
            return Err(invalid("invalid protected-range column"));
        }
        column = column
            .checked_mul(26)
            .and_then(|v| v.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| invalid("protected-range column overflow"))?;
    }
    if column == 0 || column > MAX_COLUMN {
        return Err(invalid(
            "protected-range column is outside worksheet limits",
        ));
    }
    Ok(column)
}

fn parse_row(value: &str) -> Result<u32> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid("invalid protected-range row"));
    }
    let row = value
        .parse::<u32>()
        .map_err(|_| invalid("protected-range row overflow"))?;
    if row == 0 || row > MAX_ROW {
        return Err(invalid("protected-range row is outside worksheet limits"));
    }
    Ok(row)
}

/// Writes canonical worksheet protection fragments in worksheet schema order.
///
/// Verifiers are serialized as inert metadata; this function never computes or checks a password.
pub fn write_worksheet_protection(
    metadata: &WorksheetProtectionMetadata,
    conformance: WorksheetProtectionConformance,
) -> Result<String> {
    validate_worksheet_protection_metadata(metadata)?;
    let mut xml = write_worksheet_protection_core(metadata, conformance)?;
    xml.push_str(&write_worksheet_protection_extensions(metadata, conformance)?);
    Ok(xml)
}

pub(crate) fn write_worksheet_protection_core(
    metadata: &WorksheetProtectionMetadata,
    conformance: WorksheetProtectionConformance,
) -> Result<String> {
    validate_worksheet_protection_metadata(metadata)?;
    let mut xml = String::new();
    if let Some(protection) = &metadata.sheet_protection {
        write!(
            xml,
            "<sheetProtection xmlns=\"{}\"",
            conformance.namespace()
        )
        .unwrap();
        write_verifier_attributes(&mut xml, protection.verifier.as_ref());
        write_nondefault_bool(&mut xml, "sheet", protection.sheet, false);
        write_nondefault_bool(&mut xml, "objects", protection.objects, false);
        write_nondefault_bool(&mut xml, "scenarios", protection.scenarios, false);
        write_nondefault_bool(&mut xml, "formatCells", protection.format_cells, true);
        write_nondefault_bool(&mut xml, "formatColumns", protection.format_columns, true);
        write_nondefault_bool(&mut xml, "formatRows", protection.format_rows, true);
        write_nondefault_bool(&mut xml, "insertColumns", protection.insert_columns, true);
        write_nondefault_bool(&mut xml, "insertRows", protection.insert_rows, true);
        write_nondefault_bool(
            &mut xml,
            "insertHyperlinks",
            protection.insert_hyperlinks,
            true,
        );
        write_nondefault_bool(&mut xml, "deleteColumns", protection.delete_columns, true);
        write_nondefault_bool(&mut xml, "deleteRows", protection.delete_rows, true);
        write_nondefault_bool(
            &mut xml,
            "selectLockedCells",
            protection.select_locked_cells,
            false,
        );
        write_nondefault_bool(&mut xml, "sort", protection.sort, true);
        write_nondefault_bool(&mut xml, "autoFilter", protection.auto_filter, true);
        write_nondefault_bool(&mut xml, "pivotTables", protection.pivot_tables, true);
        write_nondefault_bool(
            &mut xml,
            "selectUnlockedCells",
            protection.select_unlocked_cells,
            false,
        );
        xml.push_str("/>");
    }

    for collection in &metadata.protected_range_collections {
        match collection.source {
            ProtectedRangeSource::Core => {
                write!(
                    xml,
                    "<protectedRanges xmlns=\"{}\">",
                    conformance.namespace()
                )
                .unwrap();
                for range in &collection.ranges {
                    write_core_range(&mut xml, range)?;
                }
                xml.push_str("</protectedRanges>");
            },
            ProtectedRangeSource::Office2010 => {},
        }
    }
    Ok(xml)
}

pub(crate) fn write_worksheet_protection_extensions(
    metadata: &WorksheetProtectionMetadata,
    conformance: WorksheetProtectionConformance,
) -> Result<String> {
    validate_worksheet_protection_metadata(metadata)?;
    let mut xml = String::new();
    let office2010 = metadata
        .protected_range_collections
        .iter()
        .find(|collection| collection.source == ProtectedRangeSource::Office2010);
    if let Some(collection) = office2010 {
        write!(xml, "<extLst xmlns=\"{}\"><ext uri=\"{}\"><x14:protectedRanges xmlns:x14=\"{}\" xmlns:xm=\"{}\">",
            conformance.namespace(), PROTECTED_RANGES_EXTENSION_URI,
            std::str::from_utf8(X14).unwrap(), std::str::from_utf8(XM).unwrap()).unwrap();
        for range in &collection.ranges {
            xml.push_str("<x14:protectedRange");
            write_xml_attribute(&mut xml, "name", &range.name);
            write_verifier_attributes(&mut xml, range.verifier.as_ref());
            if let Some(descriptor) = &range.security_descriptor {
                write_xml_attribute(&mut xml, "securityDescriptor", descriptor);
            }
            xml.push_str("><xm:sqref>");
            write_xml_text(&mut xml, &canonical_sqref(&range.sqref));
            xml.push_str("</xm:sqref></x14:protectedRange>");
        }
        xml.push_str("</x14:protectedRanges></ext></extLst>");
    }
    Ok(xml)
}

pub fn validate_worksheet_protection_metadata(
    metadata: &WorksheetProtectionMetadata,
) -> Result<()> {
    if let Some(protection) = metadata.sheet_protection.as_ref() {
        validate_sheet_protection(protection)?;
    }
    let mut sources = HashSet::new();
    let mut names = HashSet::new();
    for collection in &metadata.protected_range_collections {
        if !sources.insert(collection.source) {
            return Err(invalid("duplicate protectedRanges collection"));
        }
        validate_collection(collection)?;
        for range in &collection.ranges {
            if !names.insert(range.name.to_lowercase()) {
                return Err(invalid(format!(
                    "duplicate protectedRange name '{}'",
                    range.name
                )));
            }
        }
    }
    Ok(())
}

fn validate_sheet_protection(value: &WorksheetProtection) -> Result<()> {
    if let Some(ProtectionPasswordVerifier::Strong(value)) = value.verifier.as_ref() {
        validate_strong_verifier(value)?;
    }
    Ok(())
}

fn validate_strong_verifier(value: &StrongProtectionPasswordVerifier) -> Result<()> {
    bounded_nonempty(&value.algorithm_name, "algorithmName")?;
    validate_xml_text(&value.algorithm_name, "algorithmName")?;
    if value.hash_value.is_empty() || value.hash_value.len() > MAX_BINARY_BYTES {
        return Err(invalid("hashValue has an invalid size"));
    }
    if value.salt_value.is_empty() || value.salt_value.len() > MAX_BINARY_BYTES {
        return Err(invalid("saltValue has an invalid size"));
    }
    if value.spin_count > MAX_SPIN_COUNT {
        return Err(invalid("spinCount exceeds 10000000"));
    }
    Ok(())
}

fn validate_collection(value: &WorksheetProtectedRangeCollection) -> Result<()> {
    if value.ranges.is_empty() || value.ranges.len() > MAX_RANGES {
        return Err(invalid("protectedRanges has an invalid range count"));
    }
    let mut names = HashSet::new();
    for range in &value.ranges {
        if range.source != value.source {
            return Err(invalid("protectedRange source does not match its collection"));
        }
        validate_range(range)?;
        if !names.insert(range.name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate protectedRange name '{}'",
                range.name
            )));
        }
    }
    Ok(())
}

fn validate_range(value: &WorksheetProtectedRange) -> Result<()> {
    bounded_nonempty(&value.name, "protectedRange name")?;
    validate_xml_text(&value.name, "protectedRange name")?;
    parse_sqref(&value.sqref.raw)?;
    if let Some(ProtectionPasswordVerifier::Strong(value)) = value.verifier.as_ref() {
        validate_strong_verifier(value)?;
    }
    if let Some(value) = value.security_descriptor.as_deref() {
        bounded(value, "securityDescriptor")?;
        validate_xml_text(value, "securityDescriptor")?;
    }
    Ok(())
}

fn validate_xml_text(value: &str, field: &str) -> Result<()> {
    if value.chars().any(|ch| {
        let code = ch as u32;
        !matches!(code, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
    }) {
        return Err(invalid(format!("{field} contains an invalid XML character")));
    }
    Ok(())
}

fn write_core_range(xml: &mut String, range: &WorksheetProtectedRange) -> Result<()> {
    if range.source != ProtectedRangeSource::Core {
        return Err(invalid("non-core range in core collection"));
    }
    xml.push_str("<protectedRange");
    write_xml_attribute(xml, "name", &range.name);
    write_xml_attribute(xml, "sqref", &canonical_sqref(&range.sqref));
    write_verifier_attributes(xml, range.verifier.as_ref());
    if let Some(descriptor) = &range.security_descriptor {
        write_xml_attribute(xml, "securityDescriptor", descriptor);
    }
    xml.push_str("/>");
    Ok(())
}

fn canonical_sqref(sqref: &ProtectionRangeSqref) -> String {
    sqref
        .references
        .iter()
        .map(|reference| reference.raw.as_str())
        .collect::<Vec<_>>()
        .join(" ")
}

fn write_verifier_attributes(xml: &mut String, verifier: Option<&ProtectionPasswordVerifier>) {
    match verifier {
        None => {},
        Some(ProtectionPasswordVerifier::Legacy(value)) => {
            write!(xml, " password=\"{value:04X}\"").unwrap();
        },
        Some(ProtectionPasswordVerifier::Strong(value)) => {
            write_xml_attribute(xml, "algorithmName", &value.algorithm_name);
            write_xml_attribute(xml, "hashValue", &BASE64.encode(&value.hash_value));
            write_xml_attribute(xml, "saltValue", &BASE64.encode(&value.salt_value));
            write!(xml, " spinCount=\"{}\"", value.spin_count).unwrap();
        },
    }
}

fn write_nondefault_bool(xml: &mut String, name: &str, value: bool, default: bool) {
    if value != default {
        write!(xml, " {name}=\"{}\"", if value { '1' } else { '0' }).unwrap();
    }
}

#[derive(Debug)]
struct ProtectionXmlScan {
    conformance: WorksheetProtectionConformance,
    sheet_data_end: usize,
    worksheet_close: usize,
    direct_ranges: Vec<Range<usize>>,
    x14_ranges: Vec<Range<usize>>,
    matching_ext_close: Option<usize>,
    ext_lst_close: Option<usize>,
}

/// Replace worksheet protection XML without rebuilding any unrelated worksheet content.
pub fn replace_worksheet_protection(
    worksheet_xml: &[u8],
    metadata: &WorksheetProtectionMetadata,
) -> Result<Vec<u8>> {
    let parsed = parse_worksheet_protection(worksheet_xml)?;
    validate_worksheet_protection_metadata(metadata)?;
    let scan = scan_protection_xml(worksheet_xml)?;
    let parsed_x14 = parsed
        .protected_range_collections
        .iter()
        .any(|collection| collection.source == ProtectedRangeSource::Office2010);
    if parsed.sheet_protection.is_some()
        != scan.direct_ranges.iter().any(|range| {
            worksheet_xml[range.clone()]
                .windows(b"sheetProtection".len())
                .any(|window| window == b"sheetProtection")
        })
        || parsed_x14 != !scan.x14_ranges.is_empty()
    {
        return Err(invalid(
            "worksheet protection selected through MCE cannot be mutated byte-exactly",
        ));
    }

    let direct = write_worksheet_protection_core(metadata, scan.conformance)?;
    let extensions = write_worksheet_protection_extensions(metadata, scan.conformance)?;
    let mut edits: Vec<(Range<usize>, Vec<u8>)> = scan
        .direct_ranges
        .iter()
        .chain(scan.x14_ranges.iter())
        .cloned()
        .map(|range| (range, Vec::new()))
        .collect();
    if !direct.is_empty() {
        edits.push((scan.sheet_data_end..scan.sheet_data_end, direct.into_bytes()));
    }
    if !extensions.is_empty() {
        let inner = extension_inner(&extensions)?;
        if let Some(range) = scan.x14_ranges.first() {
            if let Some(edit) = edits.iter_mut().find(|(candidate, _)| candidate == range) {
                edit.1 = inner.into_bytes();
            }
        } else if let Some(position) = scan.matching_ext_close {
            edits.push((position..position, inner.into_bytes()));
        } else if let Some(position) = scan.ext_lst_close {
            let ext = extension_wrapper(&inner, scan.conformance);
            edits.push((position..position, ext.into_bytes()));
        } else {
            edits.push((
                scan.worksheet_close..scan.worksheet_close,
                extensions.into_bytes(),
            ));
        }
    }
    apply_xml_edits(worksheet_xml, edits)
}

fn extension_inner(fragment: &str) -> Result<String> {
    let start = fragment
        .find("<x14:protectedRanges")
        .ok_or_else(|| invalid("invalid generated protection extension"))?;
    let end = fragment
        .rfind("</x14:protectedRanges>")
        .ok_or_else(|| invalid("invalid generated protection extension"))?
        + "</x14:protectedRanges>".len();
    Ok(fragment[start..end].to_string())
}

fn extension_wrapper(inner: &str, conformance: WorksheetProtectionConformance) -> String {
    format!(
        "<ext xmlns=\"{}\" uri=\"{}\">{inner}</ext>",
        conformance.namespace(),
        PROTECTED_RANGES_EXTENSION_URI
    )
}

fn apply_xml_edits(xml: &[u8], mut edits: Vec<(Range<usize>, Vec<u8>)>) -> Result<Vec<u8>> {
    edits.sort_by_key(|(range, _)| (range.start, range.end));
    let mut output = Vec::with_capacity(xml.len());
    let mut cursor = 0usize;
    for (range, replacement) in edits {
        if range.start < cursor || range.end < range.start || range.end > xml.len() {
            return Err(invalid("overlapping worksheet protection XML edits"));
        }
        output.extend_from_slice(&xml[cursor..range.start]);
        output.extend_from_slice(&replacement);
        cursor = range.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    parse_worksheet_protection(&output)?;
    Ok(output)
}

fn scan_protection_xml(xml: &[u8]) -> Result<ProtectionXmlScan> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut previous = 0usize;
    let mut conformance = None;
    let mut sheet_data_start = None;
    let mut sheet_data_end = None;
    let mut direct_start = None;
    let mut direct_ranges = Vec::new();
    let mut x14_start = None;
    let mut x14_ranges = Vec::new();
    let mut matching_ext_depth = None;
    let mut matching_ext_close = None;
    let mut ext_lst_depth = None;
    let mut ext_lst_close = None;
    let mut worksheet_close = None;
    loop {
        let start = previous;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = reader.buffer_position() as usize;
        previous = end;
        let decoder = reader.decoder();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth += 1;
                let local = element.local_name();
                if depth == 1 && local.as_ref() == b"worksheet" {
                    conformance = match namespace {
                        ResolveResult::Bound(value) if value.as_ref() == CORE => {
                            Some(WorksheetProtectionConformance::Transitional)
                        },
                        ResolveResult::Bound(value) if value.as_ref() == STRICT => {
                            Some(WorksheetProtectionConformance::Strict)
                        },
                        _ => None,
                    };
                } else if depth == 2 && spreadsheet(&namespace) {
                    match local.as_ref() {
                        b"sheetData" => sheet_data_start = Some(depth),
                        b"sheetProtection" | b"protectedRanges" => {
                            direct_start = Some((depth, start))
                        },
                        b"extLst" => ext_lst_depth = Some(depth),
                        _ => {},
                    }
                }
                if spreadsheet(&namespace)
                    && local.as_ref() == b"ext"
                    && attribute(&element, decoder, &resolver, b"uri")?.as_deref()
                        == Some(PROTECTED_RANGES_EXTENSION_URI)
                {
                    matching_ext_depth = Some(depth);
                }
                if exact(&namespace, X14)
                    && local.as_ref() == b"protectedRanges"
                    && matching_ext_depth.is_some()
                {
                    x14_start = Some((depth, start));
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                let element_depth = depth + 1;
                if element_depth == 2
                    && spreadsheet(&namespace)
                    && local.as_ref() == b"sheetData"
                {
                    sheet_data_end = Some(end);
                }
                if element_depth == 2
                    && spreadsheet(&namespace)
                    && matches!(local.as_ref(), b"sheetProtection" | b"protectedRanges")
                {
                    direct_ranges.push(start..end);
                }
                if exact(&namespace, X14)
                    && local.as_ref() == b"protectedRanges"
                    && matching_ext_depth.is_some()
                {
                    x14_ranges.push(start..end);
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                if direct_start.is_some_and(|(element_depth, _)| element_depth == depth) {
                    let (_, range_start) = direct_start.take().unwrap();
                    direct_ranges.push(range_start..end);
                }
                if x14_start.is_some_and(|(element_depth, _)| element_depth == depth) {
                    let (_, range_start) = x14_start.take().unwrap();
                    x14_ranges.push(range_start..end);
                }
                if sheet_data_start == Some(depth) {
                    sheet_data_end = Some(end);
                    sheet_data_start = None;
                }
                if matching_ext_depth == Some(depth) {
                    matching_ext_close = Some(start);
                    matching_ext_depth = None;
                }
                if ext_lst_depth == Some(depth) {
                    ext_lst_close = Some(start);
                    ext_lst_depth = None;
                }
                if depth == 1 && local.as_ref() == b"worksheet" {
                    worksheet_close = Some(start);
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(ProtectionXmlScan {
        conformance: conformance.ok_or_else(|| invalid("invalid worksheet namespace"))?,
        sheet_data_end: sheet_data_end.ok_or_else(|| invalid("worksheet is missing sheetData"))?,
        worksheet_close: worksheet_close.ok_or_else(|| invalid("worksheet is not closed"))?,
        direct_ranges,
        x14_ranges,
        matching_ext_close,
        ext_lst_close,
    })
}

fn write_xml_attribute(xml: &mut String, name: &str, value: &str) {
    write!(xml, " {name}=\"").unwrap();
    write_xml_escaped(xml, value, true);
    xml.push('"');
}

fn write_xml_text(xml: &mut String, value: &str) {
    write_xml_escaped(xml, value, false);
}

fn write_xml_escaped(xml: &mut String, value: &str, attribute: bool) {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' if attribute => xml.push_str("&quot;"),
            '\'' if attribute => xml.push_str("&apos;"),
            _ => xml.push(character),
        }
    }
}

fn decode_base64(value: &str, field: &str) -> Result<Vec<u8>> {
    let compact: String = value
        .chars()
        .filter(|character| !character.is_ascii_whitespace())
        .collect();
    if compact.len() > MAX_BINARY_BYTES.saturating_mul(2) {
        return Err(invalid(format!("{field} is too large")));
    }
    let decoded = BASE64
        .decode(compact.as_bytes())
        .map_err(|_| invalid(format!("invalid base64 in {field}")))?;
    if decoded.is_empty() || decoded.len() > MAX_BINARY_BYTES || BASE64.encode(&decoded) != compact
    {
        return Err(invalid(format!(
            "{field} is empty, non-canonical, or too large"
        )));
    }
    Ok(decoded)
}

fn attribute(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    name: &[u8],
) -> Result<Option<String>> {
    let mut value = None;
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        if namespace_declaration(attr.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attr.key);
        if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == name {
            set_once(
                &mut value,
                attr.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(xml_error)?
                    .into_owned(),
                "attribute",
            )?;
        }
    }
    Ok(value)
}

fn direct_collection_child(
    element_depth: usize,
    core: &Option<(usize, Vec<WorksheetProtectedRange>)>,
    x14: &Option<(usize, Vec<WorksheetProtectedRange>)>,
) -> bool {
    core.as_ref()
        .is_some_and(|(depth, _)| element_depth == *depth + 1)
        || x14
            .as_ref()
            .is_some_and(|(depth, _)| element_depth == *depth + 1)
}

fn update_worksheet_order(
    namespace: &ResolveResult<'_>,
    local: &[u8],
    order: &mut u8,
) -> Result<()> {
    if !spreadsheet(namespace) {
        if matches!(local, b"sheetProtection" | b"protectedRanges") {
            return Err(invalid("spoofed worksheet protection element namespace"));
        }
        return Ok(());
    }
    let rank = if matches!(
        local,
        b"sheetPr"
            | b"dimension"
            | b"sheetViews"
            | b"sheetFormatPr"
            | b"cols"
            | b"sheetData"
            | b"sheetCalcPr"
    ) {
        Some(0)
    } else if local == b"sheetProtection" {
        Some(1)
    } else if local == b"protectedRanges" {
        Some(2)
    } else if matches!(
        local,
        b"scenarios"
            | b"autoFilter"
            | b"sortState"
            | b"dataConsolidate"
            | b"customSheetViews"
            | b"mergeCells"
            | b"phoneticPr"
            | b"conditionalFormatting"
            | b"dataValidations"
            | b"hyperlinks"
            | b"printOptions"
            | b"pageMargins"
            | b"pageSetup"
            | b"headerFooter"
            | b"rowBreaks"
            | b"colBreaks"
            | b"customProperties"
            | b"cellWatches"
            | b"ignoredErrors"
            | b"smartTags"
            | b"drawing"
            | b"legacyDrawing"
            | b"legacyDrawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
    ) {
        Some(3)
    } else {
        None
    };
    if let Some(rank) = rank {
        if rank < *order {
            return Err(invalid(
                "worksheet protection family is out of schema order",
            ));
        }
        *order = (*order).max(rank);
    }
    Ok(())
}

fn namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn require_unqualified_attribute(namespace: &ResolveResult<'_>, element: &str) -> Result<()> {
    match namespace {
        ResolveResult::Unbound => Ok(()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound namespace prefix {} on {element}",
            String::from_utf8_lossy(prefix)
        ))),
        ResolveResult::Bound(_) => Err(invalid(format!("unknown namespaced {element} attribute"))),
    }
}

fn set_once(target: &mut Option<String>, value: String, field: &str) -> Result<()> {
    if target.replace(value).is_some() {
        return Err(invalid(format!("duplicate {field} attribute")));
    }
    Ok(())
}

fn parse_bool(value: &str, field: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid boolean value for {field}"))),
    }
}

fn bounded(value: &str, field: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        return Err(invalid(format!("{field} is too large")));
    }
    Ok(())
}

fn bounded_nonempty(value: &str, field: &str) -> Result<()> {
    bounded(value, field)?;
    if value.is_empty() {
        return Err(invalid(format!("{field} cannot be empty")));
    }
    Ok(())
}

fn spreadsheet(namespace: &ResolveResult<'_>) -> bool {
    exact(namespace, CORE) || exact(namespace, STRICT)
}
fn exact(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected)
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    const START: &str =
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;

    fn parse(body: &str) -> Result<WorksheetProtectionMetadata> {
        parse_worksheet_protection(format!("{START}{body}</worksheet>").as_bytes())
    }

    #[test]
    fn parses_legacy_verifier_and_schema_defaults() {
        let metadata =
            parse(r#"<sheetProtection password="CC3D" sheet="1" objects="true" scenarios="1"/>"#)
                .unwrap();
        let protection = metadata.sheet_protection().unwrap();
        assert_eq!(
            protection.verifier(),
            Some(&ProtectionPasswordVerifier::Legacy(0xCC3D))
        );
        assert!(protection.sheet_locked());
        assert!(protection.format_cells_locked());
        assert!(!protection.select_locked_cells_locked());
    }

    #[test]
    fn parses_core_strong_ranges_and_column_shorthand() {
        let metadata = parse(r#"<sheetProtection sheet="1"/><protectedRanges><protectedRange algorithmName="SHA-512" hashValue="AQI=" saltValue="AwQ=" spinCount="100000" sqref="A6 C:C" name="editable" securityDescriptor="D:test"/></protectedRanges>"#).unwrap();
        let range = metadata.protected_ranges().next().unwrap();
        assert_eq!(range.name(), "editable");
        assert_eq!(
            range.sqref().references()[0].kind(),
            ProtectionRangeReferenceKind::Cells {
                start_row: 6,
                start_column: 1,
                end_row: 6,
                end_column: 1
            }
        );
        assert_eq!(
            range.sqref().references()[1].kind(),
            ProtectionRangeReferenceKind::Columns {
                start_column: 3,
                end_column: 3
            }
        );
        let ProtectionPasswordVerifier::Strong(verifier) = range.verifier().unwrap() else {
            panic!("expected strong verifier")
        };
        assert_eq!(verifier.algorithm_name(), "SHA-512");
        assert_eq!(verifier.hash_value(), &[1, 2]);
        assert_eq!(verifier.spin_count(), 100_000);
    }

    #[test]
    fn parses_x14_extension_range() {
        let xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" mc:Ignorable="x14 xm"><extLst><ext uri="{FC87AEE6-9EDD-4A0A-B7FB-166176984837}"><x14:protectedRanges><x14:protectedRange name="Range1" password="1234"><xm:sqref>$B$2:$C$4</xm:sqref></x14:protectedRange></x14:protectedRanges></ext></extLst></worksheet>"#;
        let metadata = parse_worksheet_protection(xml).unwrap();
        let collection = &metadata.protected_range_collections()[0];
        assert_eq!(collection.source(), ProtectedRangeSource::Office2010);
        assert_eq!(
            collection.ranges()[0].sqref().references()[0].kind(),
            ProtectionRangeReferenceKind::Cells {
                start_row: 2,
                start_column: 2,
                end_row: 4,
                end_column: 3
            }
        );
    }

    #[test]
    fn rejects_incomplete_or_conflicting_verifiers() {
        assert!(parse(r#"<sheetProtection algorithmName="SHA-512" hashValue="AQI="/>"#).is_err());
        assert!(parse(r#"<protectedRanges><protectedRange name="bad" sqref="A1" password="1234" algorithmName="SHA-512" hashValue="AQI=" saltValue="AwQ=" spinCount="1"/></protectedRanges>"#).is_err());
        assert!(
            parse(
                r#"<protectedRanges><protectedRange name="bad" sqref="XFE1"/></protectedRanges>"#
            )
            .is_err()
        );
    }

    #[test]
    fn deterministic_writer_round_trips_core_x14_and_strict_metadata() {
        let source = format!(
            r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="{}" xmlns:xm="{}" mc:Ignorable="x14 xm"><sheetProtection algorithmName="SHA-512" hashValue="AQI=" saltValue="AwQ=" spinCount="0" sheet="1" formatCells="0"/><protectedRanges><protectedRange name="A&amp;B" sqref="A1  C:C" password="00af" securityDescriptor="D:&quot;test&quot;"/></protectedRanges><extLst><ext uri="{}"><x14:protectedRanges><x14:protectedRange name="X"><xm:sqref>$B$2:$C$4</xm:sqref></x14:protectedRange></x14:protectedRanges></ext></extLst></worksheet>"#,
            std::str::from_utf8(STRICT).unwrap(),
            std::str::from_utf8(X14).unwrap(),
            std::str::from_utf8(XM).unwrap(),
            PROTECTED_RANGES_EXTENSION_URI
        );
        let metadata = parse_worksheet_protection(source.as_bytes()).unwrap();
        let fragment =
            write_worksheet_protection(&metadata, WorksheetProtectionConformance::Strict).unwrap();
        assert!(fragment.contains("password=\"00AF\""));
        assert!(fragment.contains("sqref=\"A1 C:C\""));
        assert!(fragment.contains("<x14:protectedRanges"));
        let wrapped = format!(
            r#"<worksheet xmlns="{}">{fragment}</worksheet>"#,
            std::str::from_utf8(STRICT).unwrap()
        );
        let reparsed = parse_worksheet_protection(wrapped.as_bytes()).unwrap();
        assert_eq!(
            write_worksheet_protection(&reparsed, WorksheetProtectionConformance::Strict).unwrap(),
            fragment
        );
    }

    #[test]
    fn rejects_spoofed_unknown_out_of_order_and_noncanonical_metadata() {
        let invalid = [
            r#"<sheetProtection xmlns:f="urn:fake" f:sheet="1"/>"#,
            r#"<protectedRanges><protectedRange name="R" sqref="A1"/></protectedRanges><sheetProtection/>"#,
            r#"<sheetProtection/><sheetData/>"#,
            r#"<protectedRanges><unknown/></protectedRanges>"#,
            r#"<protectedRanges><protectedRange name="R" sqref="A1" algorithmName="SHA-512" hashValue="" saltValue="AQ==" spinCount="1"/></protectedRanges>"#,
            r#"<protectedRanges><protectedRange name="R" sqref="A1" algorithmName="SHA-512" hashValue="AQI=" saltValue="AQ=" spinCount="1"/></protectedRanges>"#,
            r#"<f:sheetProtection xmlns:f="urn:fake"/>"#,
        ];
        for body in invalid {
            assert!(parse(body).is_err(), "accepted {body}");
        }

        let ignored = format!(
            r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test" mc:Ignorable="x"><sheetProtection x:future="1"/></worksheet>"#,
            std::str::from_utf8(CORE).unwrap()
        );
        assert!(parse_worksheet_protection(ignored.as_bytes()).is_ok());
        let preserved = format!(
            r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test" mc:Ignorable="x" mc:PreserveAttributes="x:*"><sheetProtection x:future="1"/></worksheet>"#,
            std::str::from_utf8(CORE).unwrap()
        );
        assert!(parse_worksheet_protection(preserved.as_bytes()).is_err());
    }

    #[test]
    fn reads_poi_libreoffice_and_synthetic_package_through_worksheet_accessor() {
        use crate::xlsx::{Workbook, Worksheet, WorksheetInfo};
        use std::fs;
        use std::path::Path;
        use std::time::{SystemTime, UNIX_EPOCH};

        fn inspect(path: &Path, relationship_id: &str, expected_ranges: usize, strong_sheet: bool) {
            let workbook = Workbook::open(path).unwrap();
            let mut worksheet = Worksheet::new(
                &workbook,
                WorksheetInfo {
                    name: "Sheet1".into(),
                    relationship_id: relationship_id.into(),
                    sheet_id: 1,
                    is_active: true,
                    print_area: None,
                    repeating_rows: None,
                    repeating_columns: None,
                },
            );
            worksheet.load_data().unwrap();
            let metadata = worksheet.protection_metadata();
            assert_eq!(metadata.protected_ranges().count(), expected_ranges);
            if strong_sheet {
                assert!(matches!(
                    metadata.sheet_protection().unwrap().verifier(),
                    Some(ProtectionPasswordVerifier::Strong(_))
                ));
            }
        }

        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        inspect(
            &root.join(
                "3rdparty/poi/test-data/spreadsheet/workbookProtection-sheet_password-2013.xlsx",
            ),
            "rId1",
            0,
            true,
        );
        inspect(
            &root.join("3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/enhanced-protection.xlsx"),
            "rId1",
            5,
            false,
        );
        inspect(&root.join("3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/enhancedProtectionRangeShorthand.xlsx"), "rId2", 1, false);

        let metadata = parse(r#"<sheetData/><sheetProtection password="CC3D" sheet="1"/><protectedRanges><protectedRange name="Editable" sqref="A1:B2"/></protectedRanges>"#).unwrap();
        let fragment =
            write_worksheet_protection(&metadata, WorksheetProtectionConformance::Transitional)
                .unwrap();
        let sheet = format!(
            r#"<?xml version="1.0"?><worksheet xmlns="{}"><sheetData/>{fragment}</worksheet>"#,
            std::str::from_utf8(CORE).unwrap()
        );
        let package = make_package(&[
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", &sheet),
        ]);
        let path = std::env::temp_dir().join(format!(
            "litchi-sheet-protection-{}-{}.xlsx",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::write(&path, package).unwrap();
        inspect(&path, "rId1", 1, false);
        fs::remove_file(path).unwrap();
    }

    fn make_package(entries: &[(&str, &str)]) -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut central = Vec::new();
        for (name, value) in entries {
            let offset = bytes.len() as u32;
            let data = value.as_bytes();
            let crc = crc32(data);
            push32(&mut bytes, 0x04034b50);
            push16(&mut bytes, 20);
            push16(&mut bytes, 0);
            push16(&mut bytes, 0);
            push16(&mut bytes, 0);
            push16(&mut bytes, 0);
            push32(&mut bytes, crc);
            push32(&mut bytes, data.len() as u32);
            push32(&mut bytes, data.len() as u32);
            push16(&mut bytes, name.len() as u16);
            push16(&mut bytes, 0);
            bytes.extend_from_slice(name.as_bytes());
            bytes.extend_from_slice(data);
            push32(&mut central, 0x02014b50);
            push16(&mut central, 20);
            push16(&mut central, 20);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push32(&mut central, crc);
            push32(&mut central, data.len() as u32);
            push32(&mut central, data.len() as u32);
            push16(&mut central, name.len() as u16);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push32(&mut central, 0);
            push32(&mut central, offset);
            central.extend_from_slice(name.as_bytes());
        }
        let offset = bytes.len() as u32;
        let size = central.len() as u32;
        bytes.extend_from_slice(&central);
        push32(&mut bytes, 0x06054b50);
        push16(&mut bytes, 0);
        push16(&mut bytes, 0);
        push16(&mut bytes, entries.len() as u16);
        push16(&mut bytes, entries.len() as u16);
        push32(&mut bytes, size);
        push32(&mut bytes, offset);
        push16(&mut bytes, 0);
        bytes
    }
    fn crc32(data: &[u8]) -> u32 {
        let mut crc = !0u32;
        for byte in data {
            crc ^= u32::from(*byte);
            for _ in 0..8 {
                crc = if crc & 1 != 0 {
                    (crc >> 1) ^ 0xedb88320
                } else {
                    crc >> 1
                };
            }
        }
        !crc
    }
    fn push16(out: &mut Vec<u8>, value: u16) {
        out.extend_from_slice(&value.to_le_bytes());
    }
    fn push32(out: &mut Vec<u8>, value: u32) {
        out.extend_from_slice(&value.to_le_bytes());
    }
}
