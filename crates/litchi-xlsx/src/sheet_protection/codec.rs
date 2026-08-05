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

use crate::error::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, Limits, Name, process_markup_compatibility};

use super::model::*;

pub(crate) const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
pub(crate) const XM: &[u8] = b"http://schemas.microsoft.com/office/excel/2006/main";
pub(crate) const CORE_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const X14_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
pub(crate) const XM_NAMESPACE: &str = "http://schemas.microsoft.com/office/excel/2006/main";
pub(crate) const PROTECTED_RANGES_EXTENSION_URI: &str = "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}";
pub(crate) const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_DEPTH: usize = 128;
pub(crate) const MAX_RANGES: usize = 100_000;
pub(crate) const MAX_REFERENCES: usize = 8_192;
pub(crate) const MAX_STRING_BYTES: usize = 32 * 1024;
pub(crate) const MAX_BINARY_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_SPIN_COUNT: u32 = 10_000_000;
pub(crate) const MAX_ROW: u32 = 1_048_576;
pub(crate) const MAX_COLUMN: u32 = 16_384;

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
pub fn parse_protection(xml: &[u8]) -> Result<Metadata> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML is too large"));
    }
    let mut capabilities = Capabilities::default();
    capabilities
        .understand_namespace(String::from_utf8_lossy(X14).into_owned())
        .understand_namespace(String::from_utf8_lossy(XM).into_owned());
    capabilities.preserve_extension_element(Name {
        namespace: String::from_utf8_lossy(X14).into_owned(),
        local_name: "protectedRanges".into(),
    });
    let validated = process_markup_compatibility(xml, &capabilities, &Limits::default())?;
    parse_selected(validated.xml.as_ref())
}

fn parse_selected(xml: &[u8]) -> Result<Metadata> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut metadata = Metadata::default();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut extension_depth = None;
    let mut core_collection: Option<(usize, Vec<ProtectedRange>)> = None;
    let mut x14_collection: Option<(usize, Vec<ProtectedRange>)> = None;
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
                if let Some((sqref_depth, _)) = sqref_text.as_ref()
                    && depth > *sqref_depth
                {
                    return Err(invalid("protected-range sqref must contain only text"));
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
                        .push(ProtectedRangeCollection {
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
                        .push(ProtectedRangeCollection {
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
    core: &Option<(usize, Vec<ProtectedRange>)>,
    x14: &Option<(usize, Vec<ProtectedRange>)>,
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
    range: ProtectedRange,
    core: &mut Option<(usize, Vec<ProtectedRange>)>,
    x14: &mut Option<(usize, Vec<ProtectedRange>)>,
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
) -> Result<Protection> {
    let mut value = Protection::default();
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

fn finish_range(range: PendingRange) -> Result<ProtectedRange> {
    let sqref = range
        .sqref
        .ok_or_else(|| invalid("protectedRange is missing sqref"))?;
    Ok(ProtectedRange {
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

pub(super) fn parse_sqref(value: &str) -> Result<ProtectionRangeSqref> {
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
pub fn write_protection(metadata: &Metadata, conformance: Conformance) -> Result<String> {
    validate_metadata(metadata)?;
    let mut xml = write_core(metadata, conformance)?;
    xml.push_str(&write_extensions(metadata, conformance)?);
    Ok(xml)
}

pub fn write_core(metadata: &Metadata, conformance: Conformance) -> Result<String> {
    validate_metadata(metadata)?;
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

pub fn write_extensions(metadata: &Metadata, conformance: Conformance) -> Result<String> {
    validate_metadata(metadata)?;
    let mut xml = String::new();
    let office2010 = metadata
        .protected_range_collections
        .iter()
        .find(|collection| collection.source == ProtectedRangeSource::Office2010);
    if let Some(collection) = office2010 {
        write!(xml, "<extLst xmlns=\"{}\"><ext uri=\"{}\"><x14:protectedRanges xmlns:x14=\"{}\" xmlns:xm=\"{}\">",
            conformance.namespace(), PROTECTED_RANGES_EXTENSION_URI,
            X14_NAMESPACE, XM_NAMESPACE).unwrap();
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

pub fn validate_metadata(metadata: &Metadata) -> Result<()> {
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

pub(super) fn validate_sheet_protection(value: &Protection) -> Result<()> {
    if let Some(ProtectionPasswordVerifier::Strong(value)) = value.verifier.as_ref() {
        validate_strong_verifier(value)?;
    }
    Ok(())
}

pub(super) fn validate_strong_verifier(value: &StrongProtectionPasswordVerifier) -> Result<()> {
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

pub(super) fn validate_collection(value: &ProtectedRangeCollection) -> Result<()> {
    if value.ranges.is_empty() || value.ranges.len() > MAX_RANGES {
        return Err(invalid("protectedRanges has an invalid range count"));
    }
    let mut names = HashSet::new();
    for range in &value.ranges {
        if range.source != value.source {
            return Err(invalid(
                "protectedRange source does not match its collection",
            ));
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

pub(super) fn validate_range(value: &ProtectedRange) -> Result<()> {
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

pub(super) fn validate_xml_text(value: &str, field: &str) -> Result<()> {
    if value.chars().any(|ch| {
        let code = ch as u32;
        !matches!(code, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
    }) {
        return Err(invalid(format!(
            "{field} contains an invalid XML character"
        )));
    }
    Ok(())
}

fn write_core_range(xml: &mut String, range: &ProtectedRange) -> Result<()> {
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
    conformance: Conformance,
    sheet_data_end: usize,
    worksheet_close: usize,
    direct_ranges: Vec<Range<usize>>,
    x14_ranges: Vec<Range<usize>>,
    matching_ext_close: Option<usize>,
    ext_lst_close: Option<usize>,
}

/// Replace worksheet protection XML without rebuilding any unrelated worksheet content.
pub fn replace_protection(worksheet_xml: &[u8], metadata: &Metadata) -> Result<Vec<u8>> {
    let parsed = parse_protection(worksheet_xml)?;
    validate_metadata(metadata)?;
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
        || parsed_x14 == scan.x14_ranges.is_empty()
    {
        return Err(invalid(
            "worksheet protection selected through MCE cannot be mutated byte-exactly",
        ));
    }

    let direct = write_core(metadata, scan.conformance)?;
    let extensions = write_extensions(metadata, scan.conformance)?;
    let mut edits: Vec<(Range<usize>, Vec<u8>)> = scan
        .direct_ranges
        .iter()
        .chain(scan.x14_ranges.iter())
        .cloned()
        .map(|range| (range, Vec::new()))
        .collect();
    if !direct.is_empty() {
        edits.push((
            scan.sheet_data_end..scan.sheet_data_end,
            direct.into_bytes(),
        ));
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

fn extension_wrapper(inner: &str, conformance: Conformance) -> String {
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
    parse_protection(&output)?;
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
                            Some(Conformance::Transitional)
                        },
                        ResolveResult::Bound(value) if value.as_ref() == STRICT => {
                            Some(Conformance::Strict)
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
                if element_depth == 2 && spreadsheet(&namespace) && local.as_ref() == b"sheetData" {
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
    core: &Option<(usize, Vec<ProtectedRange>)>,
    x14: &Option<(usize, Vec<ProtectedRange>)>,
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

pub(super) fn bounded(value: &str, field: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        return Err(invalid(format!("{field} is too large")));
    }
    Ok(())
}

pub(super) fn bounded_nonempty(value: &str, field: &str) -> Result<()> {
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
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
