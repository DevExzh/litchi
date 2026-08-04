//! Typed `workbookProtection` metadata and XML codec for SpreadsheetML workbooks.
//!
//! The OOXML `workbookProtection` element carries advisory structure, window,
//! and revision locks along with optional legacy or strong password verifier
//! metadata. This module preserves and serializes that metadata but never
//! accepts or validates a password, and it never enforces an editing policy.

use std::collections::HashSet;
use std::fmt::Write as _;

use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::namespace::is_spreadsheetml_name;
use super::sheet_protection::{ProtectionPasswordVerifier, StrongProtectionPasswordVerifier};
use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::mce::process_ooxml;

const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_BINARY_BYTES: usize = 1024 * 1024;
const MAX_ENCODED_BINARY_BYTES: usize = MAX_BINARY_BYTES * 2;

/// Passive metadata from one SpreadsheetML `workbookProtection` element.
///
/// Password verifier values are opaque metadata only. They do not make a
/// workbook encrypted, and this crate neither validates passwords nor enforces
/// any of the requested locks.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorkbookProtectionMetadata {
    workbook_verifier: Option<ProtectionPasswordVerifier>,
    revisions_verifier: Option<ProtectionPasswordVerifier>,
    lock_structure: bool,
    lock_windows: bool,
    lock_revision: bool,
}

impl WorkbookProtectionMetadata {
    /// Create empty workbook-protection metadata.
    pub fn new() -> Self {
        Self::default()
    }

    /// Return the optional workbook password verifier metadata.
    pub fn workbook_verifier(&self) -> Option<&ProtectionPasswordVerifier> {
        self.workbook_verifier.as_ref()
    }

    /// Set the workbook password verifier metadata.
    pub fn set_workbook_verifier(&mut self, verifier: Option<ProtectionPasswordVerifier>) {
        self.workbook_verifier = verifier;
    }

    /// Return the optional revision password verifier metadata.
    pub fn revisions_verifier(&self) -> Option<&ProtectionPasswordVerifier> {
        self.revisions_verifier.as_ref()
    }

    /// Set the revision password verifier metadata.
    pub fn set_revisions_verifier(&mut self, verifier: Option<ProtectionPasswordVerifier>) {
        self.revisions_verifier = verifier;
    }

    /// Whether workbook structure changes are requested to be locked.
    pub fn structure_locked(&self) -> bool {
        self.lock_structure
    }

    /// Request or clear the workbook-structure lock.
    pub fn set_structure_locked(&mut self, locked: bool) {
        self.lock_structure = locked;
    }

    /// Whether workbook-window changes are requested to be locked.
    pub fn windows_locked(&self) -> bool {
        self.lock_windows
    }

    /// Request or clear the workbook-window lock.
    pub fn set_windows_locked(&mut self, locked: bool) {
        self.lock_windows = locked;
    }

    /// Whether workbook revision changes are requested to be locked.
    pub fn revision_locked(&self) -> bool {
        self.lock_revision
    }

    /// Request or clear the workbook revision-history lock.
    pub fn set_revision_locked(&mut self, locked: bool) {
        self.lock_revision = locked;
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

/// Parse an optional `workbookProtection` element from a complete `workbook.xml` part.
pub fn parse_workbook_protection(xml: &[u8]) -> Result<Option<WorkbookProtectionMetadata>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("workbook XML is too large"));
    }
    let selected = process_ooxml(xml)?;
    parse_selected(selected.as_ref())
}

/// Serialize one typed SpreadsheetML `workbookProtection` element.
///
/// The returned element uses canonical uppercase hexadecimal legacy verifiers
/// and canonical base64 for strong verifier byte strings. Passwords are never
/// accepted or derived by this codec.
pub fn write_workbook_protection(value: &WorkbookProtectionMetadata) -> Result<String> {
    let mut xml = String::with_capacity(256);
    xml.push_str("<workbookProtection");
    write_verifier(&mut xml, "workbook", value.workbook_verifier.as_ref())?;
    write_verifier(&mut xml, "revisions", value.revisions_verifier.as_ref())?;
    write_true(&mut xml, "lockStructure", value.lock_structure)?;
    write_true(&mut xml, "lockWindows", value.lock_windows)?;
    write_true(&mut xml, "lockRevision", value.lock_revision)?;
    xml.push_str("/>");
    Ok(xml)
}

fn write_verifier(
    xml: &mut String,
    prefix: &str,
    verifier: Option<&ProtectionPasswordVerifier>,
) -> Result<()> {
    match verifier {
        None => {},
        Some(ProtectionPasswordVerifier::Legacy(value)) => {
            let attribute = if prefix == "workbook" {
                "workbookPassword"
            } else {
                "revisionsPassword"
            };
            write!(xml, " {attribute}=\"{value:04X}\"").map_err(xml_error)?;
        },
        Some(ProtectionPasswordVerifier::Strong(value)) => {
            write_xml_attribute(
                xml,
                &format!("{prefix}AlgorithmName"),
                value.algorithm_name(),
            )?;
            write_xml_attribute(
                xml,
                &format!("{prefix}HashValue"),
                &BASE64.encode(value.hash_value()),
            )?;
            write_xml_attribute(
                xml,
                &format!("{prefix}SaltValue"),
                &BASE64.encode(value.salt_value()),
            )?;
            write!(xml, " {prefix}SpinCount=\"{}\"", value.spin_count()).map_err(xml_error)?;
        },
    }
    Ok(())
}

fn write_true(xml: &mut String, name: &str, value: bool) -> Result<()> {
    if value {
        write!(xml, " {name}=\"1\"").map_err(xml_error)?;
    }
    Ok(())
}

fn write_xml_attribute(xml: &mut String, name: &str, value: &str) -> Result<()> {
    if value.chars().any(|character| {
        let code = character as u32;
        !matches!(code, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
    }) {
        return Err(invalid(format!(
            "workbook protection {name} contains an invalid XML character"
        )));
    }
    write!(xml, " {name}=\"").map_err(xml_error)?;
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            _ => xml.push(character),
        }
    }
    xml.push('"');
    Ok(())
}

fn parse_selected(xml: &[u8]) -> Result<Option<WorkbookProtectionMetadata>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut protection = None;
    let mut protection_depth = None;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("workbook XML nesting overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid("workbook XML is too deep"));
                }
                if depth == 1 {
                    if root_seen
                        || root_closed
                        || !is_spreadsheetml_name(&namespace, element.name(), b"workbook")
                    {
                        return Err(invalid(
                            "workbook protection parser requires a SpreadsheetML workbook root",
                        ));
                    }
                    root_seen = true;
                    continue;
                }
                if protection_depth.is_some() {
                    return Err(invalid("workbookProtection must be empty"));
                }
                if depth == 2
                    && is_spreadsheetml_name(&namespace, element.name(), b"workbookProtection")
                {
                    if protection.is_some() {
                        return Err(invalid(
                            "workbook has duplicate workbookProtection elements",
                        ));
                    }
                    protection = Some(parse_protection_element(&element, decoder, &resolver)?);
                    protection_depth = Some(depth);
                }
            },
            Event::Empty(element) if depth == 0 => {
                if root_seen
                    || root_closed
                    || !is_spreadsheetml_name(&namespace, element.name(), b"workbook")
                {
                    return Err(invalid(
                        "workbook protection parser requires a SpreadsheetML workbook root",
                    ));
                }
                root_seen = true;
                root_closed = true;
            },
            Event::Empty(element) => {
                if protection_depth.is_some() {
                    return Err(invalid("workbookProtection must be empty"));
                }
                if depth == 1
                    && is_spreadsheetml_name(&namespace, element.name(), b"workbookProtection")
                {
                    if protection.is_some() {
                        return Err(invalid(
                            "workbook has duplicate workbookProtection elements",
                        ));
                    }
                    protection = Some(parse_protection_element(&element, decoder, &resolver)?);
                }
            },
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_)
                if protection_depth.is_some() =>
            {
                return Err(invalid("workbookProtection must be empty"));
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("workbook XML closes an element outside its root"));
                }
                if protection_depth == Some(depth) {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"workbookProtection") {
                        return Err(invalid("workbookProtection has an invalid closing element"));
                    }
                    protection_depth = None;
                }
                if depth == 1 {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"workbook") {
                        return Err(invalid("workbook XML has an invalid root closing element"));
                    }
                    root_closed = true;
                }
                depth -= 1;
            },
            Event::Eof
                if !root_seen || !root_closed || depth != 0 || protection_depth.is_some() =>
            {
                return Err(invalid("workbook XML is incomplete"));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    Ok(protection)
}

fn parse_protection_element(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<WorkbookProtectionMetadata> {
    let mut value = WorkbookProtectionMetadata::default();
    let mut workbook = RawCredential::default();
    let mut revisions = RawCredential::default();
    let mut seen = HashSet::new();

    for attribute in element.attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        require_unqualified_attribute(&namespace, "workbookProtection")?;
        if !seen.insert(local.as_ref().to_vec()) {
            return Err(invalid("duplicate workbookProtection attribute"));
        }
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if text.len() > MAX_ENCODED_BINARY_BYTES {
            return Err(invalid("workbookProtection attribute is too large"));
        }
        match local.as_ref() {
            b"workbookPassword" => workbook.password = Some(text),
            b"workbookAlgorithmName" => workbook.algorithm_name = Some(text),
            b"workbookHashValue" => workbook.hash_value = Some(text),
            b"workbookSaltValue" => workbook.salt_value = Some(text),
            b"workbookSpinCount" => workbook.spin_count = Some(text),
            b"revisionsPassword" => revisions.password = Some(text),
            b"revisionsAlgorithmName" => revisions.algorithm_name = Some(text),
            b"revisionsHashValue" => revisions.hash_value = Some(text),
            b"revisionsSaltValue" => revisions.salt_value = Some(text),
            b"revisionsSpinCount" => revisions.spin_count = Some(text),
            b"lockStructure" => value.lock_structure = parse_bool(&text, "lockStructure")?,
            b"lockWindows" => value.lock_windows = parse_bool(&text, "lockWindows")?,
            b"lockRevision" => value.lock_revision = parse_bool(&text, "lockRevision")?,
            other => {
                return Err(invalid(format!(
                    "unknown workbookProtection attribute '{}'",
                    String::from_utf8_lossy(other)
                )));
            },
        }
    }

    value.workbook_verifier = finish_credential(workbook, "workbook")?;
    value.revisions_verifier = finish_credential(revisions, "revisions")?;
    Ok(value)
}

fn finish_credential(
    raw: RawCredential,
    prefix: &str,
) -> Result<Option<ProtectionPasswordVerifier>> {
    let strong_present = raw.algorithm_name.is_some()
        || raw.hash_value.is_some()
        || raw.salt_value.is_some()
        || raw.spin_count.is_some();
    if raw.password.is_some() && strong_present {
        return Err(invalid(format!(
            "{prefix} legacy password and strong hash metadata are mutually exclusive"
        )));
    }
    if let Some(password) = raw.password {
        if password.len() != 4 || !password.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(invalid(format!(
                "{prefix} password verifier must be four hexadecimal digits"
            )));
        }
        return Ok(Some(ProtectionPasswordVerifier::Legacy(
            u16::from_str_radix(&password, 16)
                .map_err(|_| invalid(format!("invalid {prefix} password verifier")))?,
        )));
    }
    if !strong_present {
        return Ok(None);
    }

    let algorithm_name = raw
        .algorithm_name
        .ok_or_else(|| invalid(format!("{prefix} strong verifier is missing algorithmName")))?;
    let hash_value = decode_base64(
        &raw.hash_value
            .ok_or_else(|| invalid(format!("{prefix} strong verifier is missing hashValue")))?,
        &format!("{prefix} hashValue"),
    )?;
    let salt_value = decode_base64(
        &raw.salt_value
            .ok_or_else(|| invalid(format!("{prefix} strong verifier is missing saltValue")))?,
        &format!("{prefix} saltValue"),
    )?;
    let spin_count = raw
        .spin_count
        .ok_or_else(|| invalid(format!("{prefix} strong verifier is missing spinCount")))?
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid {prefix} spinCount")))?;

    StrongProtectionPasswordVerifier::new(algorithm_name, hash_value, salt_value, spin_count)
        .map(ProtectionPasswordVerifier::Strong)
        .map(Some)
        .map_err(OoxmlError::from)
}

fn decode_base64(value: &str, field: &str) -> Result<Vec<u8>> {
    let compact: String = value
        .chars()
        .filter(|character| !character.is_ascii_whitespace())
        .collect();
    if compact.len() > MAX_ENCODED_BINARY_BYTES {
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

fn parse_bool(value: &str, field: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid boolean value for {field}"))),
    }
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

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::Workbook;
    use std::path::Path;

    const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    fn workbook(body: &str) -> String {
        format!(r#"<workbook xmlns="{TRANSITIONAL}">{body}</workbook>"#)
    }

    #[test]
    fn parses_legacy_workbook_and_revision_verifiers() {
        let xml = workbook(
            r#"<workbookProtection workbookPassword="aBcD" revisionsPassword="0011" lockStructure="true" lockWindows="1" lockRevision="false"/>"#,
        );
        let metadata = parse_workbook_protection(xml.as_bytes()).unwrap().unwrap();
        assert_eq!(
            metadata.workbook_verifier(),
            Some(&ProtectionPasswordVerifier::Legacy(0xABCD))
        );
        assert_eq!(
            metadata.revisions_verifier(),
            Some(&ProtectionPasswordVerifier::Legacy(0x0011))
        );
        assert!(metadata.structure_locked());
        assert!(metadata.windows_locked());
        assert!(!metadata.revision_locked());
    }

    #[test]
    fn parses_strict_strong_workbook_protection() {
        let xml = format!(
            r#"<workbook xmlns="{STRICT}"><workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="AQI=" workbookSaltValue="AwQ=" workbookSpinCount="100000" lockRevision="1"/></workbook>"#
        );
        let metadata = parse_workbook_protection(xml.as_bytes()).unwrap().unwrap();
        let Some(ProtectionPasswordVerifier::Strong(verifier)) = metadata.workbook_verifier()
        else {
            panic!("expected strong workbook verifier");
        };
        assert_eq!(verifier.algorithm_name(), "SHA-512");
        assert_eq!(verifier.hash_value(), [1, 2]);
        assert_eq!(verifier.salt_value(), [3, 4]);
        assert_eq!(verifier.spin_count(), 100_000);
        assert!(metadata.revision_locked());
    }

    #[test]
    fn rejects_invalid_or_mixed_verifier_metadata() {
        let mixed = workbook(
            r#"<workbookProtection workbookPassword="ABCD" workbookAlgorithmName="SHA-512" workbookHashValue="AQI=" workbookSaltValue="AwQ=" workbookSpinCount="1"/>"#,
        );
        assert!(parse_workbook_protection(mixed.as_bytes()).is_err());

        let incomplete = workbook(
            r#"<workbookProtection revisionsAlgorithmName="SHA-512" revisionsHashValue="AQI=" revisionsSaltValue="AwQ="/>"#,
        );
        assert!(parse_workbook_protection(incomplete.as_bytes()).is_err());
    }

    #[test]
    fn writes_and_reparses_all_typed_protection_attributes() {
        let mut metadata = WorkbookProtectionMetadata::new();
        metadata.set_workbook_verifier(Some(ProtectionPasswordVerifier::Strong(
            StrongProtectionPasswordVerifier::new("SHA-512", vec![1, 2, 3], vec![4, 5, 6], 100_000)
                .unwrap(),
        )));
        metadata.set_revisions_verifier(Some(ProtectionPasswordVerifier::Legacy(0x00AF)));
        metadata.set_structure_locked(true);
        metadata.set_windows_locked(true);
        metadata.set_revision_locked(true);

        let element = write_workbook_protection(&metadata).unwrap();
        assert_eq!(
            element,
            r#"<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="AQID" workbookSaltValue="BAUG" workbookSpinCount="100000" revisionsPassword="00AF" lockStructure="1" lockWindows="1" lockRevision="1"/>"#
        );

        let parsed = parse_workbook_protection(workbook(&element).as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(parsed, metadata);
    }

    #[test]
    fn writer_escapes_future_algorithm_names_without_weakening_validation() {
        let mut metadata = WorkbookProtectionMetadata::new();
        metadata.set_revisions_verifier(Some(ProtectionPasswordVerifier::Strong(
            StrongProtectionPasswordVerifier::new("future&hash", vec![1], vec![2], 0).unwrap(),
        )));

        let element = write_workbook_protection(&metadata).unwrap();
        assert!(element.contains(r#"revisionsAlgorithmName="future&amp;hash""#));
        let parsed = parse_workbook_protection(workbook(&element).as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(parsed, metadata);
    }

    #[test]
    fn workbook_accessor_reads_a_real_excel_strong_verifier() {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let workbook = Workbook::open(root.join(
            "test-data/poi/test-data/spreadsheet/workbookProtection-workbook_password-2013.xlsx",
        ))
        .unwrap();
        let metadata = workbook.workbook_protection_metadata().unwrap().unwrap();
        assert!(metadata.structure_locked());
        assert!(!metadata.windows_locked());
        assert!(!metadata.revision_locked());
        let Some(ProtectionPasswordVerifier::Strong(verifier)) = metadata.workbook_verifier()
        else {
            panic!("expected strong workbook verifier");
        };
        assert_eq!(verifier.algorithm_name(), "SHA-512");
        assert_eq!(verifier.spin_count(), 100_000);
        assert!(metadata.revisions_verifier().is_none());
    }
}
