//! Immutable worksheet-protection metadata for SpreadsheetML worksheets.

use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use quick_xml::XmlVersion;

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

/// Source schema for a protected-range collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
    pub fn algorithm_name(&self) -> &str { &self.algorithm_name }
    pub fn hash_value(&self) -> &[u8] { &self.hash_value }
    pub fn salt_value(&self) -> &[u8] { &self.salt_value }
    pub fn spin_count(&self) -> u32 { self.spin_count }
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
    pub fn verifier(&self) -> Option<&ProtectionPasswordVerifier> { self.verifier.as_ref() }
    pub fn sheet_locked(&self) -> bool { self.sheet }
    pub fn objects_locked(&self) -> bool { self.objects }
    pub fn scenarios_locked(&self) -> bool { self.scenarios }
    pub fn format_cells_locked(&self) -> bool { self.format_cells }
    pub fn format_columns_locked(&self) -> bool { self.format_columns }
    pub fn format_rows_locked(&self) -> bool { self.format_rows }
    pub fn insert_columns_locked(&self) -> bool { self.insert_columns }
    pub fn insert_rows_locked(&self) -> bool { self.insert_rows }
    pub fn insert_hyperlinks_locked(&self) -> bool { self.insert_hyperlinks }
    pub fn delete_columns_locked(&self) -> bool { self.delete_columns }
    pub fn delete_rows_locked(&self) -> bool { self.delete_rows }
    pub fn select_locked_cells_locked(&self) -> bool { self.select_locked_cells }
    pub fn sort_locked(&self) -> bool { self.sort }
    pub fn auto_filter_locked(&self) -> bool { self.auto_filter }
    pub fn pivot_tables_locked(&self) -> bool { self.pivot_tables }
    pub fn select_unlocked_cells_locked(&self) -> bool { self.select_unlocked_cells }
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
    Cells { start_row: u32, start_column: u32, end_row: u32, end_column: u32 },
    Columns { start_column: u32, end_column: u32 },
    Rows { start_row: u32, end_row: u32 },
}

/// One validated reference in a protected range's `sqref`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionRangeReference {
    raw: String,
    kind: ProtectionRangeReferenceKind,
}

impl ProtectionRangeReference {
    pub fn raw(&self) -> &str { &self.raw }
    pub fn kind(&self) -> ProtectionRangeReferenceKind { self.kind }
}

/// Validated whitespace-separated protected-range references.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionRangeSqref {
    raw: String,
    references: Vec<ProtectionRangeReference>,
}

impl ProtectionRangeSqref {
    pub fn raw(&self) -> &str { &self.raw }
    pub fn references(&self) -> &[ProtectionRangeReference] { &self.references }
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
    pub fn source(&self) -> ProtectedRangeSource { self.source }
    pub fn name(&self) -> &str { &self.name }
    pub fn sqref(&self) -> &ProtectionRangeSqref { &self.sqref }
    pub fn verifier(&self) -> Option<&ProtectionPasswordVerifier> { self.verifier.as_ref() }
    pub fn security_descriptor(&self) -> Option<&str> { self.security_descriptor.as_deref() }
}

/// A protected-range container in worksheet document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetProtectedRangeCollection {
    source: ProtectedRangeSource,
    ranges: Vec<WorksheetProtectedRange>,
}

impl WorksheetProtectedRangeCollection {
    pub fn source(&self) -> ProtectedRangeSource { self.source }
    pub fn ranges(&self) -> &[WorksheetProtectedRange] { &self.ranges }
}

/// Complete worksheet protection metadata.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorksheetProtectionMetadata {
    sheet_protection: Option<WorksheetProtection>,
    protected_range_collections: Vec<WorksheetProtectedRangeCollection>,
}

impl WorksheetProtectionMetadata {
    pub fn sheet_protection(&self) -> Option<&WorksheetProtection> { self.sheet_protection.as_ref() }
    pub fn protected_range_collections(&self) -> &[WorksheetProtectedRangeCollection] { &self.protected_range_collections }
    pub fn protected_ranges(&self) -> impl Iterator<Item = &WorksheetProtectedRange> {
        self.protected_range_collections.iter().flat_map(|collection| collection.ranges.iter())
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
    if xml.len() > MAX_XML_BYTES { return Err(invalid("worksheet XML is too large")); }
    let mut capabilities = MceCapabilities::default();
    capabilities.understand_namespace(String::from_utf8_lossy(X14).into_owned())
        .understand_namespace(String::from_utf8_lossy(XM).into_owned());
    capabilities.preserve_extension_element(ExpandedName {
        namespace: String::from_utf8_lossy(X14).into_owned(),
        local_name: "protectedRanges".into(),
    });
    let validated = process_markup_compatibility(xml, &capabilities, &MceLimits::default())?;
    let selected = if validated.report.alternate_content_count == 0 { xml } else { validated.xml.as_ref() };
    parse_selected(selected)
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

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                if depth > MAX_DEPTH { return Err(invalid("worksheet XML nesting is too deep")); }
                if depth == 1 {
                    if root_seen || !spreadsheet(&namespace) || element.local_name().as_ref() != b"worksheet" {
                        return Err(invalid("worksheet protection parser requires a worksheet root"));
                    }
                    root_seen = true;
                    continue;
                }
                if sheet_protection_depth.is_some() { return Err(invalid("sheetProtection must be empty")); }
                if let Some((sqref_depth, _)) = sqref_text.as_ref() {
                    if depth > *sqref_depth { return Err(invalid("protected-range sqref must contain only text")); }
                }
                if spreadsheet(&namespace) && element.local_name().as_ref() == b"ext" {
                    if attribute(&element, decoder, b"uri")?.as_deref() == Some(PROTECTED_RANGES_EXTENSION_URI) {
                        extension_depth = Some(depth);
                    }
                } else if depth == 2 && spreadsheet(&namespace) && element.local_name().as_ref() == b"sheetProtection" {
                    if metadata.sheet_protection.is_some() { return Err(invalid("duplicate sheetProtection element")); }
                    metadata.sheet_protection = Some(parse_sheet_protection(&element, decoder)?);
                    sheet_protection_depth = Some(depth);
                } else if depth == 2 && spreadsheet(&namespace) && element.local_name().as_ref() == b"protectedRanges" {
                    if core_collection.is_some() || metadata.protected_range_collections.iter().any(|c| c.source == ProtectedRangeSource::Core) {
                        return Err(invalid("duplicate core protectedRanges element"));
                    }
                    core_collection = Some((depth, Vec::new()));
                } else if exact(&namespace, X14) && element.local_name().as_ref() == b"protectedRanges" && extension_depth.is_some() {
                    if x14_collection.is_some() { return Err(invalid("nested x14 protectedRanges element")); }
                    x14_collection = Some((depth, Vec::new()));
                } else if element.local_name().as_ref() == b"protectedRange" {
                    let source = collection_source(depth, &namespace, &core_collection, &x14_collection)?;
                    if pending.is_some() { return Err(invalid("nested protectedRange element")); }
                    pending = Some((depth, parse_pending_range(&element, decoder, source)?));
                } else if exact(&namespace, XM)
                    && element.local_name().as_ref() == b"sqref"
                    && x14_collection.is_some()
                {
                    let Some((range_depth, range)) = pending.as_ref() else { return Err(invalid("x14 sqref is outside protectedRange")); };
                    if range.source != ProtectedRangeSource::Office2010 || depth != *range_depth + 1 || sqref_text.is_some() {
                        return Err(invalid("invalid x14 protected-range sqref placement"));
                    }
                    sqref_text = Some((depth, String::new()));
                } else if pending.is_some() {
                    return Err(invalid("unexpected child in protectedRange"));
                }
            }
            Event::Empty(element) => {
                if sheet_protection_depth.is_some() { return Err(invalid("sheetProtection must be empty")); }
                if depth == 1 && spreadsheet(&namespace) && element.local_name().as_ref() == b"sheetProtection" {
                    if metadata.sheet_protection.is_some() { return Err(invalid("duplicate sheetProtection element")); }
                    metadata.sheet_protection = Some(parse_sheet_protection(&element, decoder)?);
                } else if element.local_name().as_ref() == b"protectedRange" {
                    let source = collection_source(depth + 1, &namespace, &core_collection, &x14_collection)?;
                    let range = finish_range(parse_pending_range(&element, decoder, source)?)?;
                    push_range(range, &mut core_collection, &mut x14_collection)?;
                } else if exact(&namespace, XM)
                    && element.local_name().as_ref() == b"sqref"
                    && x14_collection.is_some()
                {
                    if pending.is_some() {
                        return Err(invalid("x14 protected-range sqref cannot be empty"));
                    }
                    return Err(invalid("x14 sqref is outside protectedRange"));
                }
            }
            Event::Text(text) => {
                if let Some((_, value)) = sqref_text.as_mut() {
                    let decoded = text.decode().map_err(xml_error)?;
                    if value.len().saturating_add(decoded.len()) > MAX_STRING_BYTES { return Err(invalid("protected-range sqref is too large")); }
                    value.push_str(&decoded);
                } else if (sheet_protection_depth.is_some() || pending.is_some()) && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("protection elements cannot contain text"));
                }
            }
            Event::End(element) => {
                if sqref_text.as_ref().is_some_and(|(sqref_depth, _)| *sqref_depth == depth) {
                    let (_, value) = sqref_text.take().expect("checked above");
                    let Some((_, range)) = pending.as_mut() else { return Err(invalid("orphan x14 sqref")); };
                    if range.sqref.replace(value).is_some() { return Err(invalid("duplicate x14 protected-range sqref")); }
                } else if pending.as_ref().is_some_and(|(range_depth, _)| *range_depth == depth) {
                    let (_, range) = pending.take().expect("checked above");
                    push_range(finish_range(range)?, &mut core_collection, &mut x14_collection)?;
                } else if core_collection.as_ref().is_some_and(|(collection_depth, _)| *collection_depth == depth) {
                    let (_, ranges) = core_collection.take().expect("checked above");
                    if ranges.is_empty() { return Err(invalid("protectedRanges must contain at least one protectedRange")); }
                    metadata.protected_range_collections.push(WorksheetProtectedRangeCollection { source: ProtectedRangeSource::Core, ranges });
                } else if x14_collection.as_ref().is_some_and(|(collection_depth, _)| *collection_depth == depth) {
                    let (_, ranges) = x14_collection.take().expect("checked above");
                    if ranges.is_empty() { return Err(invalid("x14 protectedRanges must contain at least one protectedRange")); }
                    metadata.protected_range_collections.push(WorksheetProtectedRangeCollection { source: ProtectedRangeSource::Office2010, ranges });
                }
                if sheet_protection_depth == Some(depth) { sheet_protection_depth = None; }
                if extension_depth == Some(depth) { extension_depth = None; }
                if depth == 1 {
                    if !spreadsheet(&namespace) || element.local_name().as_ref() != b"worksheet" { return Err(invalid("invalid worksheet closing element")); }
                    root_closed = true;
                }
                depth = depth.checked_sub(1).ok_or_else(|| invalid("unexpected XML end element"))?;
            }
            Event::CData(_) => return Err(invalid("CDATA is not allowed in worksheet protection metadata")),
            Event::DocType(_) | Event::PI(_) => return Err(invalid("DTD and processing instructions are rejected")),
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") { return Err(invalid("custom XML entities are rejected")); }
            }
            Event::Decl(_) | Event::Comment(_) => {}
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || pending.is_some() || core_collection.is_some() || x14_collection.is_some() {
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
    if core.as_ref().is_some_and(|(depth, _)| element_depth == *depth + 1) && spreadsheet(namespace) {
        Ok(ProtectedRangeSource::Core)
    } else if x14.as_ref().is_some_and(|(depth, _)| element_depth == *depth + 1) && exact(namespace, X14) {
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
        ProtectedRangeSource::Core => &mut core.as_mut().ok_or_else(|| invalid("missing core protectedRanges parent"))?.1,
        ProtectedRangeSource::Office2010 => &mut x14.as_mut().ok_or_else(|| invalid("missing x14 protectedRanges parent"))?.1,
    };
    if ranges.len() >= MAX_RANGES { return Err(invalid("too many protected ranges")); }
    ranges.push(range);
    Ok(())
}

fn parse_sheet_protection(element: &BytesStart<'_>, decoder: Decoder) -> Result<WorksheetProtection> {
    let mut value = WorksheetProtection::default();
    let mut credential = RawCredential::default();
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        if attr.key.as_ref().contains(&b':') { continue; }
        let text = attr
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        match attr.key.local_name().as_ref() {
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
            b"selectLockedCells" => value.select_locked_cells = parse_bool(&text, "selectLockedCells")?,
            b"sort" => value.sort = parse_bool(&text, "sort")?,
            b"autoFilter" => value.auto_filter = parse_bool(&text, "autoFilter")?,
            b"pivotTables" => value.pivot_tables = parse_bool(&text, "pivotTables")?,
            b"selectUnlockedCells" => value.select_unlocked_cells = parse_bool(&text, "selectUnlockedCells")?,
            other => return Err(invalid(format!("unknown sheetProtection attribute '{}'", String::from_utf8_lossy(other)))),
        }
    }
    value.verifier = finish_credential(credential)?;
    Ok(value)
}

fn parse_pending_range(element: &BytesStart<'_>, decoder: Decoder, source: ProtectedRangeSource) -> Result<PendingRange> {
    let mut name = None;
    let mut sqref = None;
    let mut security_descriptor = None;
    let mut credential = RawCredential::default();
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        if attr.key.as_ref().contains(&b':') { continue; }
        let text = attr
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        match attr.key.local_name().as_ref() {
            b"name" => set_once(&mut name, text, "name")?,
            b"sqref" if source == ProtectedRangeSource::Core => set_once(&mut sqref, text, "sqref")?,
            b"password" => set_once(&mut credential.password, text, "password")?,
            b"algorithmName" => set_once(&mut credential.algorithm_name, text, "algorithmName")?,
            b"hashValue" => set_once(&mut credential.hash_value, text, "hashValue")?,
            b"saltValue" => set_once(&mut credential.salt_value, text, "saltValue")?,
            b"spinCount" => set_once(&mut credential.spin_count, text, "spinCount")?,
            b"securityDescriptor" => set_once(&mut security_descriptor, text, "securityDescriptor")?,
            other => return Err(invalid(format!("unknown protectedRange attribute '{}'", String::from_utf8_lossy(other)))),
        }
    }
    let name = name.ok_or_else(|| invalid("protectedRange is missing name"))?;
    bounded_nonempty(&name, "protectedRange name")?;
    if let Some(value) = security_descriptor.as_ref() { bounded(value, "securityDescriptor")?; }
    Ok(PendingRange { source, name, sqref, credential, security_descriptor })
}

fn finish_range(range: PendingRange) -> Result<WorksheetProtectedRange> {
    let sqref = range.sqref.ok_or_else(|| invalid("protectedRange is missing sqref"))?;
    Ok(WorksheetProtectedRange {
        source: range.source,
        name: range.name,
        sqref: parse_sqref(&sqref)?,
        verifier: finish_credential(range.credential)?,
        security_descriptor: range.security_descriptor,
    })
}

fn finish_credential(raw: RawCredential) -> Result<Option<ProtectionPasswordVerifier>> {
    let strong_present = raw.algorithm_name.is_some() || raw.hash_value.is_some() || raw.salt_value.is_some() || raw.spin_count.is_some();
    if raw.password.is_some() && strong_present { return Err(invalid("legacy password and strong hash metadata are mutually exclusive")); }
    if let Some(password) = raw.password {
        if password.len() != 4 || !password.bytes().all(|byte| byte.is_ascii_hexdigit()) { return Err(invalid("legacy password verifier must be four hexadecimal digits")); }
        return Ok(Some(ProtectionPasswordVerifier::Legacy(u16::from_str_radix(&password, 16).map_err(|_| invalid("invalid legacy password verifier"))?)));
    }
    if !strong_present { return Ok(None); }
    let algorithm_name = raw.algorithm_name.ok_or_else(|| invalid("strong verifier is missing algorithmName"))?;
    bounded_nonempty(&algorithm_name, "algorithmName")?;
    let hash_value = decode_base64(&raw.hash_value.ok_or_else(|| invalid("strong verifier is missing hashValue"))?, "hashValue")?;
    let salt_value = decode_base64(&raw.salt_value.ok_or_else(|| invalid("strong verifier is missing saltValue"))?, "saltValue")?;
    let spin_count = raw.spin_count.ok_or_else(|| invalid("strong verifier is missing spinCount"))?.parse::<u32>().map_err(|_| invalid("invalid spinCount"))?;
    if spin_count > MAX_SPIN_COUNT { return Err(invalid("spinCount exceeds 10000000")); }
    Ok(Some(ProtectionPasswordVerifier::Strong(StrongProtectionPasswordVerifier { algorithm_name, hash_value, salt_value, spin_count })))
}

fn parse_sqref(value: &str) -> Result<ProtectionRangeSqref> {
    bounded_nonempty(value, "protectedRange sqref")?;
    let tokens: Vec<_> = value.split_ascii_whitespace().collect();
    if tokens.is_empty() || tokens.len() > MAX_REFERENCES { return Err(invalid("invalid protectedRange sqref reference count")); }
    let references = tokens.into_iter().map(parse_reference).collect::<Result<Vec<_>>>()?;
    Ok(ProtectionRangeSqref { raw: value.to_string(), references })
}

fn parse_reference(raw: &str) -> Result<ProtectionRangeReference> {
    let parts: Vec<_> = raw.split(':').collect();
    let kind = match parts.as_slice() {
        [single] => match parse_endpoint(single)? {
            Endpoint::Cell(row, column) => ProtectionRangeReferenceKind::Cells { start_row: row, start_column: column, end_row: row, end_column: column },
            _ => return Err(invalid("single protected-range reference must be a cell")),
        },
        [start, end] => match (parse_endpoint(start)?, parse_endpoint(end)?) {
            (Endpoint::Cell(sr, sc), Endpoint::Cell(er, ec)) if sr <= er && sc <= ec => ProtectionRangeReferenceKind::Cells { start_row: sr, start_column: sc, end_row: er, end_column: ec },
            (Endpoint::Column(sc), Endpoint::Column(ec)) if sc <= ec => ProtectionRangeReferenceKind::Columns { start_column: sc, end_column: ec },
            (Endpoint::Row(sr), Endpoint::Row(er)) if sr <= er => ProtectionRangeReferenceKind::Rows { start_row: sr, end_row: er },
            _ => return Err(invalid("invalid or reversed protected-range reference")),
        },
        _ => return Err(invalid("protected-range reference contains too many colons")),
    };
    Ok(ProtectionRangeReference { raw: raw.to_string(), kind })
}

enum Endpoint { Cell(u32, u32), Column(u32), Row(u32) }

fn parse_endpoint(value: &str) -> Result<Endpoint> {
    let value = value.strip_prefix('$').unwrap_or(value);
    if value.is_empty() { return Err(invalid("empty protected-range endpoint")); }
    let letter_count = value.bytes().take_while(u8::is_ascii_alphabetic).count();
    if letter_count > 0 {
        let column_text = &value[..letter_count];
        let rest = value[letter_count..].strip_prefix('$').unwrap_or(&value[letter_count..]);
        let column = parse_column(column_text)?;
        if rest.is_empty() { return Ok(Endpoint::Column(column)); }
        let row = parse_row(rest)?;
        Ok(Endpoint::Cell(row, column))
    } else {
        Ok(Endpoint::Row(parse_row(value)?))
    }
}

fn parse_column(value: &str) -> Result<u32> {
    let mut column = 0u32;
    for byte in value.bytes() {
        if !byte.is_ascii_alphabetic() { return Err(invalid("invalid protected-range column")); }
        column = column.checked_mul(26).and_then(|v| v.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))).ok_or_else(|| invalid("protected-range column overflow"))?;
    }
    if column == 0 || column > MAX_COLUMN { return Err(invalid("protected-range column is outside worksheet limits")); }
    Ok(column)
}

fn parse_row(value: &str) -> Result<u32> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) { return Err(invalid("invalid protected-range row")); }
    let row = value.parse::<u32>().map_err(|_| invalid("protected-range row overflow"))?;
    if row == 0 || row > MAX_ROW { return Err(invalid("protected-range row is outside worksheet limits")); }
    Ok(row)
}

fn decode_base64(value: &str, field: &str) -> Result<Vec<u8>> {
    let compact: String = value.chars().filter(|character| !character.is_ascii_whitespace()).collect();
    if compact.len() > MAX_BINARY_BYTES.saturating_mul(2) { return Err(invalid(format!("{field} is too large"))); }
    let decoded = BASE64.decode(compact.as_bytes()).map_err(|_| invalid(format!("invalid base64 in {field}")))?;
    if decoded.len() > MAX_BINARY_BYTES { return Err(invalid(format!("{field} is too large"))); }
    Ok(decoded)
}

fn attribute(element: &BytesStart<'_>, decoder: Decoder, name: &[u8]) -> Result<Option<String>> {
    let mut value = None;
    for attr in element.attributes() {
        let attr = attr.map_err(xml_error)?;
        if !attr.key.as_ref().contains(&b':') && attr.key.local_name().as_ref() == name {
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

fn set_once(target: &mut Option<String>, value: String, field: &str) -> Result<()> {
    if target.replace(value).is_some() { return Err(invalid(format!("duplicate {field} attribute"))); }
    Ok(())
}

fn parse_bool(value: &str, field: &str) -> Result<bool> {
    match value { "1" | "true" => Ok(true), "0" | "false" => Ok(false), _ => Err(invalid(format!("invalid boolean value for {field}"))) }
}

fn bounded(value: &str, field: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES { return Err(invalid(format!("{field} is too large"))); }
    Ok(())
}

fn bounded_nonempty(value: &str, field: &str) -> Result<()> {
    bounded(value, field)?;
    if value.is_empty() { return Err(invalid(format!("{field} cannot be empty"))); }
    Ok(())
}

fn spreadsheet(namespace: &ResolveResult<'_>) -> bool { exact(namespace, CORE) || exact(namespace, STRICT) }
fn exact(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool { matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected) }
fn invalid(message: impl Into<String>) -> OoxmlError { OoxmlError::InvalidFormat(message.into()) }
fn xml_error(error: impl std::fmt::Display) -> OoxmlError { OoxmlError::Xml(error.to_string()) }

#[cfg(test)]
mod tests {
    use super::*;

    const START: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;

    fn parse(body: &str) -> Result<WorksheetProtectionMetadata> {
        parse_worksheet_protection(format!("{START}{body}</worksheet>").as_bytes())
    }

    #[test]
    fn parses_legacy_verifier_and_schema_defaults() {
        let metadata = parse(r#"<sheetProtection password="CC3D" sheet="1" objects="true" scenarios="1"/>"#).unwrap();
        let protection = metadata.sheet_protection().unwrap();
        assert_eq!(protection.verifier(), Some(&ProtectionPasswordVerifier::Legacy(0xCC3D)));
        assert!(protection.sheet_locked());
        assert!(protection.format_cells_locked());
        assert!(!protection.select_locked_cells_locked());
    }

    #[test]
    fn parses_core_strong_ranges_and_column_shorthand() {
        let metadata = parse(r#"<sheetProtection sheet="1"/><protectedRanges><protectedRange algorithmName="SHA-512" hashValue="AQI=" saltValue="AwQ=" spinCount="100000" sqref="A6 C:C" name="editable" securityDescriptor="D:test"/></protectedRanges>"#).unwrap();
        let range = metadata.protected_ranges().next().unwrap();
        assert_eq!(range.name(), "editable");
        assert_eq!(range.sqref().references()[0].kind(), ProtectionRangeReferenceKind::Cells { start_row: 6, start_column: 1, end_row: 6, end_column: 1 });
        assert_eq!(range.sqref().references()[1].kind(), ProtectionRangeReferenceKind::Columns { start_column: 3, end_column: 3 });
        let ProtectionPasswordVerifier::Strong(verifier) = range.verifier().unwrap() else { panic!("expected strong verifier") };
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
        assert_eq!(collection.ranges()[0].sqref().references()[0].kind(), ProtectionRangeReferenceKind::Cells { start_row: 2, start_column: 2, end_row: 4, end_column: 3 });
    }

    #[test]
    fn rejects_incomplete_or_conflicting_verifiers() {
        assert!(parse(r#"<sheetProtection algorithmName="SHA-512" hashValue="AQI="/>"#).is_err());
        assert!(parse(r#"<protectedRanges><protectedRange name="bad" sqref="A1" password="1234" algorithmName="SHA-512" hashValue="AQI=" saltValue="AwQ=" spinCount="1"/></protectedRanges>"#).is_err());
        assert!(parse(r#"<protectedRanges><protectedRange name="bad" sqref="XFE1"/></protectedRanges>"#).is_err());
    }
}
