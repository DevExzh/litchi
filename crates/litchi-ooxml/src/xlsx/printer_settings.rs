//! SpreadsheetML worksheet Printer Settings references and inert DEVMODE parts.

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use crate::xlsx::page_setup::parse_complete_worksheet_page_setup;
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const PRINTER_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
const STRICT_PRINTER_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/printerSettings";
const PRINTER_CT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_SETTINGS_BYTES: usize = 16 * 1024 * 1024;
const MAX_RELATIONSHIP_ID_BYTES: usize = 4096;
const MAX_DEPTH: usize = 256;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PrinterSettingsConformance { Transitional, Strict }

impl PrinterSettingsConformance {
    fn sml(self) -> &'static str { match self { Self::Transitional => SML, Self::Strict => STRICT_SML } }
    fn rel(self) -> &'static str { match self { Self::Transitional => REL, Self::Strict => STRICT_REL } }
    fn printer_rel(self) -> &'static str { match self { Self::Transitional => PRINTER_REL, Self::Strict => STRICT_PRINTER_REL } }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetPrinterSettingsReference { pub relationship_id: String }

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrinterSettingsResource {
    pub part_name: String,
    /// Complete DEVMODE structure, including driver-private bytes, retained inertly.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetPrinterSettings {
    pub reference: WorksheetPrinterSettingsReference,
    pub resource: PrinterSettingsResource,
}

/// Parses the optional Printer Settings relationship reference from a worksheet.
pub fn parse_worksheet_printer_settings_reference(xml: &[u8]) -> Result<Option<WorksheetPrinterSettingsReference>> {
    if xml.len() > MAX_XML_BYTES { return Err(limit("worksheet XML bytes")); }
    validate_mce(xml)?;
    let setup = parse_complete_worksheet_page_setup(xml)?;
    let Some(id) = setup.as_ref().and_then(|value| value.printer_settings_relationship_id()) else { return Ok(None); };
    validate_id(id)?;
    Ok(Some(WorksheetPrinterSettingsReference { relationship_id: id.to_owned() }))
}

/// Deterministically writes a self-contained `pageSetup` reference fragment.
pub fn write_worksheet_printer_settings_reference(reference: &WorksheetPrinterSettingsReference, conformance: PrinterSettingsConformance) -> Result<Vec<u8>> {
    validate_id(&reference.relationship_id)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<x:pageSetup xmlns:x=\""); escape(&mut output, conformance.sml());
    output.extend_from_slice(b"\" xmlns:r=\""); escape(&mut output, conformance.rel());
    output.extend_from_slice(b"\" r:id=\""); escape(&mut output, &reference.relationship_id);
    output.extend_from_slice(b"\"/>");
    Ok(output)
}

/// Resolves and validates the Printer Settings part for one worksheet.
pub fn load_worksheet_printer_settings(package: &OpcPackage, worksheet_name: &PackURI) -> Result<Option<WorksheetPrinterSettings>> {
    if package.rels().iter().any(|relationship| is_printer_relationship(relationship.reltype())) { return Err(invalid("package root cannot source a Printer Settings relationship")); }
    let worksheet = package.get_part(worksheet_name)?;
    require_worksheet(worksheet)?;
    let reference = parse_worksheet_printer_settings_reference(worksheet.blob())?;
    let mut relationships = worksheet.rels().iter().filter(|relationship| is_printer_relationship(relationship.reltype()));
    let relationship = relationships.next();
    if relationships.next().is_some() { return Err(invalid("worksheet has multiple Printer Settings relationships")); }
    match (reference, relationship) {
        (None, None) => Ok(None),
        (None, Some(_)) => Err(invalid("worksheet has an unreferenced Printer Settings relationship")),
        (Some(_), None) => Err(invalid("pageSetup references a missing Printer Settings relationship")),
        (Some(reference), Some(relationship)) => {
            if relationship.r_id() != reference.relationship_id { return Err(invalid(format!("pageSetup references '{}', but Printer Settings relationship is '{}'", reference.relationship_id, relationship.r_id()))); }
            if relationship.is_external() { return Err(invalid("Printer Settings relationship must be internal")); }
            let target = relationship.target_partname()?;
            if !target.as_str().starts_with("/xl/printerSettings/") { return Err(invalid(format!("Printer Settings part '{target}' is outside /xl/printerSettings"))); }
            let part = package.get_part(&target)?;
            if part.content_type() != PRINTER_CT { return Err(invalid(format!("Printer Settings part '{target}' has invalid content type '{}'", part.content_type()))); }
            if !part.rels().is_empty() { return Err(invalid(format!("Printer Settings part '{target}' has forbidden outbound relationships"))); }
            validate_settings_bytes(part.blob())?;
            Ok(Some(WorksheetPrinterSettings { reference, resource: PrinterSettingsResource { part_name: target.to_string(), data: part.blob().to_vec() } }))
        }
    }
}

/// Adds one inert Printer Settings part and its `pageSetup` reference to a worksheet.
pub fn store_worksheet_printer_settings(package: &mut OpcPackage, worksheet_name: &PackURI, value: &WorksheetPrinterSettings, conformance: PrinterSettingsConformance) -> Result<()> {
    validate_id(&value.reference.relationship_id)?;
    validate_settings_bytes(&value.resource.data)?;
    let resource_uri = PackURI::new(&value.resource.part_name).map_err(OoxmlError::InvalidUri)?;
    if !resource_uri.as_str().starts_with("/xl/printerSettings/") { return Err(invalid(format!("Printer Settings part '{resource_uri}' is outside /xl/printerSettings"))); }
    if load_worksheet_printer_settings(package, worksheet_name)?.is_some() { return Err(invalid("worksheet already has Printer Settings")); }
    let worksheet = package.get_part(worksheet_name)?;
    require_worksheet(worksheet)?;
    if worksheet.rels().get(&value.reference.relationship_id).is_some() { return Err(invalid(format!("worksheet relationship ID '{}' already exists", value.reference.relationship_id))); }
    if package.iter_parts().any(|part| part.partname() == &resource_uri) { return Err(invalid(format!("Printer Settings part '{resource_uri}' already exists"))); }
    let actual = worksheet_conformance(worksheet.blob())?;
    if actual != conformance { return Err(invalid("requested conformance does not match worksheet namespace")); }
    let updated = add_reference_to_worksheet(worksheet.blob(), &value.reference, conformance)?;
    let target = resource_uri.relative_ref(worksheet_name.base_uri());
    package.get_part_mut(worksheet_name)?.set_blob(updated);
    package.add_part(Box::new(BlobPart::new(resource_uri, PRINTER_CT.into(), value.resource.data.clone())));
    package.get_part_mut(worksheet_name)?.rels_mut().add_relationship(conformance.printer_rel().into(), target, value.reference.relationship_id.clone(), false);
    Ok(())
}

fn add_reference_to_worksheet(xml: &[u8], reference: &WorksheetPrinterSettingsReference, conformance: PrinterSettingsConformance) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(xml);
    let later = [b"headerFooter".as_slice(), b"rowBreaks", b"colBreaks", b"customProperties", b"cellWatches", b"ignoredErrors", b"smartTags", b"drawing", b"legacyDrawing", b"legacyDrawingHF", b"picture", b"oleObjects", b"controls", b"webPublishItems", b"tableParts", b"extLst"];
    let mut depth = 0usize; let mut root = false; let mut insert = None; let mut page_setup = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_| invalid("worksheet XML offset overflow"))?;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let end = usize::try_from(reader.buffer_position()).map_err(|_| invalid("worksheet XML offset overflow"))?;
        match event {
            Event::Start(element) => {
                let core = exact(&namespace, conformance.sml());
                if depth == 0 { if !core || element.local_name().as_ref() != b"worksheet" { return Err(invalid("worksheet root does not match requested conformance")); } root = true; }
                else if depth == 1 && core && element.local_name().as_ref() == b"pageSetup" { if page_setup.replace((end - 1, false)).is_some() { return Err(invalid("worksheet has multiple direct pageSetup elements")); } }
                else if depth == 1 && core && later.contains(&element.local_name().as_ref()) { insert.get_or_insert(start); }
                depth = depth.checked_add(1).ok_or_else(|| limit("worksheet XML depth"))?; if depth > MAX_DEPTH { return Err(limit("worksheet XML depth")); }
            }
            Event::Empty(element) => {
                let core = exact(&namespace, conformance.sml());
                if depth == 1 && core && element.local_name().as_ref() == b"pageSetup" { if page_setup.replace((end - 2, true)).is_some() { return Err(invalid("worksheet has multiple direct pageSetup elements")); } }
                else if depth == 1 && core && later.contains(&element.local_name().as_ref()) { insert.get_or_insert(start); }
            }
            Event::End(element) => { if depth == 0 { return Err(invalid("unexpected worksheet closing element")); } if depth == 1 && element.local_name().as_ref() == b"worksheet" { insert.get_or_insert(start); } depth -= 1; }
            Event::DocType(_) | Event::PI(_) => return Err(invalid("DTDs and processing instructions are rejected")),
            Event::Eof => break,
            _ => {}
        }
    }
    if !root || depth != 0 { return Err(invalid("invalid worksheet XML")); }
    let mut addition = Vec::new();
    if let Some((position, empty)) = page_setup {
        addition.extend_from_slice(b" xmlns:r=\""); escape(&mut addition, conformance.rel()); addition.extend_from_slice(b"\" r:id=\""); escape(&mut addition, &reference.relationship_id); addition.push(b'\"');
        let mut output = Vec::with_capacity(xml.len() + addition.len()); output.extend_from_slice(&xml[..position]); output.extend_from_slice(&addition); output.extend_from_slice(&xml[position..]);
        if empty && !output[position + addition.len()..].starts_with(b"/>") { return Err(invalid("invalid empty pageSetup serialization")); }
        return Ok(output);
    }
    let position = insert.ok_or_else(|| invalid("missing worksheet closing element"))?;
    addition = write_worksheet_printer_settings_reference(reference, conformance)?;
    let length = xml.len().checked_add(addition.len()).ok_or_else(|| limit("updated worksheet XML bytes"))?; if length > MAX_XML_BYTES { return Err(limit("updated worksheet XML bytes")); }
    let mut output = Vec::with_capacity(length); output.extend_from_slice(&xml[..position]); output.extend_from_slice(&addition); output.extend_from_slice(&xml[position..]); Ok(output)
}

fn validate_mce(xml: &[u8]) -> Result<()> {
    let limits = MceLimits { max_input_bytes: MAX_XML_BYTES, max_output_bytes: MAX_XML_BYTES, max_depth: MAX_DEPTH, max_namespace_bindings: 4096, max_directive_tokens: 4096, max_choices_per_alternate: 1024 };
    process_markup_compatibility(xml, &MceCapabilities::default(), &limits)?;
    Ok(())
}

fn worksheet_conformance(xml: &[u8]) -> Result<PrinterSettingsConformance> {
    if xml.len() > MAX_XML_BYTES { return Err(limit("worksheet XML bytes")); }
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) | Event::Empty(element) if element.local_name().as_ref() == b"worksheet" => return match namespace { ResolveResult::Bound(Namespace(value)) if value.as_ref() == SML.as_bytes() => Ok(PrinterSettingsConformance::Transitional), ResolveResult::Bound(Namespace(value)) if value.as_ref() == STRICT_SML.as_bytes() => Ok(PrinterSettingsConformance::Strict), _ => Err(invalid("unsupported worksheet namespace")) },
            Event::DocType(_) | Event::PI(_) => return Err(invalid("DTDs and processing instructions are rejected")),
            Event::Eof => return Err(invalid("missing worksheet root")), _ => {}
        }
    }
}

fn require_worksheet(part: &dyn Part) -> Result<()> { if part.content_type() == ct::SML_WORKSHEET { Ok(()) } else { Err(invalid(format!("part '{}' is not a worksheet", part.partname()))) } }
fn validate_settings_bytes(data: &[u8]) -> Result<()> { if data.is_empty() { Err(invalid("Printer Settings DEVMODE bytes cannot be empty")) } else if data.len() > MAX_SETTINGS_BYTES { Err(limit("DEVMODE bytes")) } else { Ok(()) } }
fn validate_id(id: &str) -> Result<()> { if id.is_empty() || id.len() > MAX_RELATIONSHIP_ID_BYTES { return Err(invalid("invalid Printer Settings relationship ID length")); } let mut bytes = id.bytes(); let first = bytes.next().unwrap(); if !(first.is_ascii_alphabetic() || first == b'_') || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')) { Err(invalid(format!("invalid Printer Settings relationship ID '{id}'"))) } else { Ok(()) } }
fn is_printer_relationship(value: &str) -> bool { matches!(value, PRINTER_REL | STRICT_PRINTER_REL) }
fn exact(namespace: &ResolveResult<'_>, value: &str) -> bool { matches!(namespace, ResolveResult::Bound(Namespace(namespace)) if { let bytes: &[u8] = namespace.as_ref(); bytes == value.as_bytes() }) }
fn escape(output: &mut Vec<u8>, value: &str) { for character in value.chars() { match character { '&' => output.extend_from_slice(b"&amp;"), '<' => output.extend_from_slice(b"&lt;"), '"' => output.extend_from_slice(b"&quot;"), '\t' => output.extend_from_slice(b"&#x9;"), '\n' => output.extend_from_slice(b"&#xA;"), '\r' => output.extend_from_slice(b"&#xD;"), _ => { let mut bytes = [0; 4]; output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes()); } } } }
fn xml_error(error: impl std::fmt::Display) -> OoxmlError { OoxmlError::Xml(error.to_string()) }
fn invalid(message: impl Into<String>) -> OoxmlError { OoxmlError::InvalidFormat(message.into()) }
fn limit(name: &str) -> OoxmlError { invalid(format!("worksheet Printer Settings {name} limit exceeded")) }

#[cfg(test)]
mod tests {
    use super::*;

    fn root() -> std::path::PathBuf { std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..") }
    fn package(conformance: PrinterSettingsConformance, page_setup: &str) -> (OpcPackage, PackURI) { let mut package = OpcPackage::new(); let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap(); package.add_part(Box::new(BlobPart::new(uri.clone(), ct::SML_WORKSHEET.into(), format!("<x:worksheet xmlns:x=\"{}\"><x:sheetData/>{page_setup}<x:headerFooter/></x:worksheet>", conformance.sml()).into_bytes()))); (package, uri) }
    fn value() -> WorksheetPrinterSettings { WorksheetPrinterSettings { reference: WorksheetPrinterSettingsReference { relationship_id: "rIdPrinter".into() }, resource: PrinterSettingsResource { part_name: "/xl/printerSettings/printerSettings1.bin".into(), data: vec![0x44, 0x45, 0x56, 0x4d, 0x4f, 0x44, 0x45] } } }

    #[test]
    fn strict_reference_round_trip_and_mce_fallback() {
        let reference = WorksheetPrinterSettingsReference { relationship_id: "rId9".into() }; let fragment = write_worksheet_printer_settings_reference(&reference, PrinterSettingsConformance::Strict).unwrap(); let xml = [format!("<x:worksheet xmlns:x=\"{STRICT_SML}\">" ).as_bytes(), fragment.as_slice(), b"</x:worksheet>"].concat(); assert_eq!(parse_worksheet_printer_settings_reference(&xml).unwrap().unwrap(), reference);
        let mce = format!(r#"<x:worksheet xmlns:x="{SML}" xmlns:r="{REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:future"><mc:AlternateContent><mc:Choice Requires="f"><x:pageSetup r:id="future"/></mc:Choice><mc:Fallback><x:pageSetup r:id="fallback"/></mc:Fallback></mc:AlternateContent></x:worksheet>"#); assert_eq!(parse_worksheet_printer_settings_reference(mce.as_bytes()).unwrap().unwrap().relationship_id, "fallback");
    }

    #[test]
    fn loads_poi_and_libreoffice_reference_packages() {
        let poi = OpcPackage::open(root().join("3rdparty/poi/test-data/spreadsheet/sample.xlsx")).unwrap(); let settings = load_worksheet_printer_settings(&poi, &PackURI::new("/xl/worksheets/sheet1.xml").unwrap()).unwrap().unwrap(); assert_eq!(settings.reference.relationship_id, "rId1"); assert_eq!(settings.resource.data.len(), 2452);
        let libreoffice = OpcPackage::open(root().join("3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf136721_letter_sized_paper.xlsx")).unwrap(); let settings = load_worksheet_printer_settings(&libreoffice, &PackURI::new("/xl/worksheets/sheet1.xml").unwrap()).unwrap().unwrap(); assert_eq!(settings.reference.relationship_id, "rId1"); assert!(!settings.resource.data.is_empty());
    }

    #[test]
    fn strict_package_writer_preserves_page_setup_and_schema_order() {
        let (mut package, uri) = package(PrinterSettingsConformance::Strict, "<x:pageSetup orientation=\"landscape\"/>"); let expected = value(); store_worksheet_printer_settings(&mut package, &uri, &expected, PrinterSettingsConformance::Strict).unwrap(); assert_eq!(load_worksheet_printer_settings(&package, &uri).unwrap().unwrap(), expected); let xml = package.get_part(&uri).unwrap().blob(); let text = std::str::from_utf8(xml).unwrap(); assert!(text.contains("orientation=\"landscape\"")); assert!(text.contains("r:id=\"rIdPrinter\"")); assert!(text.find("pageSetup").unwrap() < text.find("headerFooter").unwrap());
    }

    #[test]
    fn package_writer_inserts_missing_page_setup_deterministically() {
        let (mut package, uri) = package(PrinterSettingsConformance::Transitional, ""); store_worksheet_printer_settings(&mut package, &uri, &value(), PrinterSettingsConformance::Transitional).unwrap(); let xml = std::str::from_utf8(package.get_part(&uri).unwrap().blob()).unwrap(); assert!(xml.find("pageSetup").unwrap() < xml.find("headerFooter").unwrap());
    }

    #[test]
    fn rejects_malformed_references_caps_and_graphs() {
        for xml in [format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><pageSetup r:id="9bad"/></worksheet>"#), format!(r#"<worksheet xmlns="{SML}"><pageSetup><x/></pageSetup></worksheet>"#), format!(r#"<!DOCTYPE x><worksheet xmlns="{SML}"/>"#)] { assert!(parse_worksheet_printer_settings_reference(xml.as_bytes()).is_err(), "{xml}"); }
        assert!(parse_worksheet_printer_settings_reference(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
        let (mut external, uri) = package(PrinterSettingsConformance::Transitional, &format!("<x:pageSetup xmlns:r=\"{REL}\" r:id=\"rId1\"/>")); external.get_part_mut(&uri).unwrap().rels_mut().add_relationship(PRINTER_REL.into(), "https://invalid.example/settings".into(), "rId1".into(), true); assert!(load_worksheet_printer_settings(&external, &uri).is_err());
        let (mut outbound, uri) = package(PrinterSettingsConformance::Transitional, ""); store_worksheet_printer_settings(&mut outbound, &uri, &value(), PrinterSettingsConformance::Transitional).unwrap(); outbound.get_part_mut(&PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap()).unwrap().rels_mut().add_relationship("urn:forbidden".into(), "x".into(), "rId1".into(), false); assert!(load_worksheet_printer_settings(&outbound, &uri).is_err());
        let mut too_large = value(); too_large.resource.data = vec![0; MAX_SETTINGS_BYTES + 1]; let (mut package, uri) = package(PrinterSettingsConformance::Transitional, ""); assert!(store_worksheet_printer_settings(&mut package, &uri, &too_large, PrinterSettingsConformance::Transitional).is_err());
    }
}
