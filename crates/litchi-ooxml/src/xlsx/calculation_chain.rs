//! Typed, inert SpreadsheetML calculation-chain metadata.

use crate::error::{OoxmlError, Result};
use crate::xlsx::Cell;
use litchi_core::sheet::Result as SheetResult;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const TRANSITIONAL_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_NS: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
const RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
const STRICT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain";
const MAX_CELLS: usize = 2_000_000;
const MAX_EXTENSION_BYTES: usize = 16 * 1024 * 1024;
const MAX_EXTENSION_ATTRIBUTES: usize = 256;
const MAX_EXTENSION_DEPTH: usize = 128;
const MAX_REFERENCE_BYTES: usize = 32;

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

/// Namespace family used by the calculation-chain writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CalculationChainConformance {
    #[default]
    Transitional,
    Strict,
}

impl CalculationChainConformance {
    const fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_NS,
            Self::Strict => STRICT_NS,
        }
    }

    const fn relationship_type(self) -> &'static str {
        match self {
            Self::Transitional => RELATIONSHIP,
            Self::Strict => STRICT_RELATIONSHIP,
        }
    }
}

/// An MCE-preserved, non-schema attribute retained without interpretation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CalculationChainExtensionAttribute {
    qualified_name: String,
    value: String,
}

impl CalculationChainExtensionAttribute {
    pub fn qualified_name(&self) -> &str {
        &self.qualified_name
    }
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// One formula cell in calculation order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CalculationCell {
    reference: String,
    sheet_id: Option<i32>,
    child_chain: Option<bool>,
    new_dependency_level: bool,
    new_thread: bool,
    array_formula: bool,
    extension_attributes: Vec<CalculationChainExtensionAttribute>,
}

impl CalculationCell {
    pub fn new(reference: impl Into<String>) -> Result<Self> {
        let reference = reference.into();
        validate_reference(&reference)?;
        Ok(Self {
            reference,
            sheet_id: None,
            child_chain: None,
            new_dependency_level: false,
            new_thread: false,
            array_formula: false,
            extension_attributes: Vec::new(),
        })
    }

    pub fn reference(&self) -> &str {
        &self.reference
    }
    pub fn sheet_id(&self) -> Option<i32> {
        self.sheet_id
    }
    pub fn child_chain_override(&self) -> Option<bool> {
        self.child_chain
    }
    pub fn starts_new_dependency_level(&self) -> bool {
        self.new_dependency_level
    }
    pub fn starts_new_thread(&self) -> bool {
        self.new_thread
    }
    pub fn is_array_formula(&self) -> bool {
        self.array_formula
    }
    pub fn extension_attributes(&self) -> &[CalculationChainExtensionAttribute] {
        &self.extension_attributes
    }

    pub fn set_sheet_id(&mut self, value: Option<i32>) -> &mut Self {
        self.sheet_id = value;
        self
    }
    pub fn set_child_chain_override(&mut self, value: Option<bool>) -> &mut Self {
        self.child_chain = value;
        self
    }
    pub fn set_starts_new_dependency_level(&mut self, value: bool) -> &mut Self {
        self.new_dependency_level = value;
        self
    }
    pub fn set_starts_new_thread(&mut self, value: bool) -> &mut Self {
        self.new_thread = value;
        self
    }
    pub fn set_array_formula(&mut self, value: bool) -> &mut Self {
        self.array_formula = value;
        self
    }
}

/// Ordered metadata from the workbook's single Calculation Chain part.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CalculationChain {
    cells: Vec<CalculationCell>,
    extension_list_xml: Option<String>,
    namespace_declarations: Vec<(String, String)>,
    extension_attributes: Vec<CalculationChainExtensionAttribute>,
}

impl CalculationChain {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn cells(&self) -> &[CalculationCell] {
        &self.cells
    }
    pub fn cells_mut(&mut self) -> &mut Vec<CalculationCell> {
        &mut self.cells
    }
    pub fn push(&mut self, cell: CalculationCell) -> Result<&mut Self> {
        if self.cells.len() >= MAX_CELLS {
            return Err(invalid("calculation chain has too many cells"));
        }
        self.cells.push(cell);
        Ok(self)
    }
    pub fn extension_list_xml(&self) -> Option<&str> {
        self.extension_list_xml.as_deref()
    }
    pub fn extension_attributes(&self) -> &[CalculationChainExtensionAttribute] {
        &self.extension_attributes
    }

    /// Resolve the inherited sheet ID at `index`, if any preceding record specifies one.
    pub fn effective_sheet_id(&self, index: usize) -> Option<i32> {
        self.cells
            .get(..=index)?
            .iter()
            .rev()
            .find_map(|cell| cell.sheet_id)
    }

    /// Resolve the inherited child-chain flag at `index` (false before the first override).
    pub fn effective_child_chain(&self, index: usize) -> Option<bool> {
        self.cells.get(..=index).map(|cells| {
            cells
                .iter()
                .rev()
                .find_map(|cell| cell.child_chain)
                .unwrap_or(false)
        })
    }

    pub fn to_xml(&self, conformance: CalculationChainConformance) -> Result<String> {
        if self.cells.is_empty() {
            return Err(invalid("calculation chain must contain at least one cell"));
        }
        if self.cells.len() > MAX_CELLS {
            return Err(invalid("calculation chain has too many cells"));
        }
        let mut xml =
            String::with_capacity(self.cells.len().saturating_mul(32).saturating_add(256));
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str("<calcChain xmlns=\"");
        xml.push_str(conformance.namespace());
        xml.push('"');
        for (name, value) in &self.namespace_declarations {
            if name != "xmlns" {
                xml.push(' ');
                xml.push_str(name);
                xml.push_str("=\"");
                escape_attribute(&mut xml, value);
                xml.push('"');
            }
        }
        write_extension_attributes(&mut xml, &self.extension_attributes)?;
        xml.push('>');
        for cell in &self.cells {
            validate_reference(&cell.reference)?;
            xml.push_str("<c r=\"");
            escape_attribute(&mut xml, &cell.reference);
            xml.push('"');
            if let Some(value) = cell.sheet_id {
                xml.push_str(" i=\"");
                xml.push_str(&value.to_string());
                xml.push('"');
            }
            if let Some(value) = cell.child_chain {
                write_bool_attribute(&mut xml, "s", value);
            }
            if cell.new_dependency_level {
                write_bool_attribute(&mut xml, "l", true);
            }
            if cell.new_thread {
                write_bool_attribute(&mut xml, "t", true);
            }
            if cell.array_formula {
                write_bool_attribute(&mut xml, "a", true);
            }
            write_extension_attributes(&mut xml, &cell.extension_attributes)?;
            xml.push_str("/>");
        }
        if let Some(extension) = &self.extension_list_xml {
            if extension.len() > MAX_EXTENSION_BYTES {
                return Err(invalid("calculation-chain extension list is too large"));
            }
            xml.push_str(extension);
        }
        xml.push_str("</calcChain>");
        Ok(xml)
    }
}

/// Parse an isolated Calculation Chain part. Formula text is never evaluated.
pub fn parse_calculation_chain(xml: &[u8]) -> Result<CalculationChain> {
    let processed = crate::common::mce::process_ooxml(xml)
        .map_err(|error| invalid(format!("calculation-chain MCE error: {error}")))?;
    let bytes = processed.as_ref();
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut chain = CalculationChain::new();
    let mut saw_root = false;
    let mut closed_root = false;
    let mut saw_extensions = false;
    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if !saw_root => {
                validate_root(&namespace, &element, closed_root)?;
                saw_root = true;
                parse_root_attributes(&element, decoder, &resolver, &mut chain)?;
            },
            Event::Empty(element) if !saw_root => {
                validate_root(&namespace, &element, closed_root)?;
                saw_root = true;
                closed_root = true;
                parse_root_attributes(&element, decoder, &resolver, &mut chain)?;
            },
            Event::Empty(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"c") =>
            {
                if saw_extensions {
                    return Err(invalid("calculation cells must precede extLst"));
                }
                push_cell(&mut chain, parse_cell(&element, decoder, &resolver)?)?;
            },
            Event::Start(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"c") =>
            {
                if saw_extensions {
                    return Err(invalid("calculation cells must precede extLst"));
                }
                let cell = parse_cell(&element, decoder, &resolver)?;
                consume_leaf(&mut reader, b"c")?;
                push_cell(&mut chain, cell)?;
            },
            Event::Empty(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"extLst") =>
            {
                if std::mem::replace(&mut saw_extensions, true) {
                    return Err(invalid("duplicate calculation-chain extLst"));
                }
                let end = position(&reader)?;
                chain.extension_list_xml = Some(raw_range(bytes, start, end)?);
            },
            Event::Start(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"extLst") =>
            {
                if std::mem::replace(&mut saw_extensions, true) {
                    return Err(invalid("duplicate calculation-chain extLst"));
                }
                let end = consume_extension_list(&mut reader)?;
                chain.extension_list_xml = Some(raw_range(bytes, start, end)?);
            },
            Event::Start(element) | Event::Empty(element) if saw_root && !closed_root => {
                return Err(invalid(format!(
                    "unexpected calculation-chain child '{}'",
                    String::from_utf8_lossy(element.local_name().as_ref())
                )));
            },
            Event::End(element)
                if saw_root && !closed_root && element.local_name().as_ref() == b"calcChain" =>
            {
                closed_root = true
            },
            Event::Text(text)
                if text
                    .decode()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                    .trim()
                    .is_empty() => {},
            Event::Comment(_) | Event::Decl(_) => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DTD and processing instructions are rejected in calculation-chain XML",
                ));
            },
            Event::Eof => break,
            _ => return Err(invalid("invalid calculation-chain XML structure")),
        }
    }
    if !saw_root || !closed_root {
        return Err(invalid("calculation-chain XML has no complete root"));
    }
    if chain.cells.is_empty() {
        return Err(invalid("calculation chain must contain at least one cell"));
    }
    Ok(chain)
}

/// Load the optional inert calculation-chain part selected by the package workbook.
///
/// Formula cells are parsed as metadata only; no formula is evaluated.
pub fn load_calculation_chain_from_package(
    package: &OpcPackage,
) -> Result<Option<CalculationChain>> {
    let workbook_uri = main_workbook_uri(package)?;
    load_calculation_chain_for_workbook(package, &workbook_uri)
}

/// Store a caller-authored inert calculation chain in a SpreadsheetML package.
///
/// The supplied order is serialized without recalculating formulas or inferring
/// dependencies. Existing calculation-chain graph violations are rejected
/// before any package part is changed. The requested conformance is applied to
/// both the part XML and its workbook relationship.
pub fn store_calculation_chain(
    package: &mut OpcPackage,
    chain: &CalculationChain,
    conformance: CalculationChainConformance,
) -> Result<()> {
    let xml = chain.to_xml(conformance)?.into_bytes();
    let workbook_uri = main_workbook_uri(package)?;
    let existing = calculation_chain_relationship(package, &workbook_uri)?;

    if let Some(existing) = existing {
        validate_calculation_chain_part_set(package, Some(&existing.part_name))?;
        validate_calculation_chain_part(package, &existing.part_name)?;
        package.get_part_mut(&existing.part_name)?.set_blob(xml);
        if existing.conformance != conformance {
            let workbook = package.get_part_mut(&workbook_uri)?;
            workbook.rels_mut().remove(&existing.relationship_id);
            workbook.rels_mut().add_relationship(
                conformance.relationship_type().into(),
                existing.target_reference,
                existing.relationship_id,
                false,
            );
        }
    } else {
        validate_calculation_chain_part_set(package, None)?;
        let part_name = next_calculation_chain_part_name(package)?;
        let relationship_id = next_calculation_chain_relationship_id(package, &workbook_uri)?;
        let target = part_name.relative_ref(workbook_uri.base_uri());
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name,
            CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(&workbook_uri)?
            .rels_mut()
            .add_relationship(
                conformance.relationship_type().into(),
                target,
                relationship_id,
                false,
            );
    }

    let _ = package.clear_digital_signatures();
    Ok(())
}

/// Remove the workbook's calculation-chain relationship and its unreferenced part.
///
/// No formulas are changed. A target that is also referenced elsewhere in the
/// package is retained.
pub fn remove_calculation_chain(package: &mut OpcPackage) -> Result<bool> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(existing) = calculation_chain_relationship(package, &workbook_uri)? else {
        validate_calculation_chain_part_set(package, None)?;
        return Ok(false);
    };
    validate_calculation_chain_part_set(package, Some(&existing.part_name))?;
    validate_calculation_chain_part(package, &existing.part_name)?;

    package
        .get_part_mut(&workbook_uri)?
        .rels_mut()
        .remove(&existing.relationship_id);
    if !package_part_is_referenced(package, &existing.part_name) {
        package.remove_part(&existing.part_name);
    }
    let _ = package.clear_digital_signatures();
    Ok(true)
}

pub(crate) fn load_calculation_chain(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> SheetResult<Option<(CalculationChain, CalculationChainConformance)>> {
    load_calculation_chain_with_conformance_for_workbook(package, workbook_uri).map_err(Into::into)
}

fn load_calculation_chain_for_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<CalculationChain>> {
    Ok(
        load_calculation_chain_with_conformance_for_workbook(package, workbook_uri)?
            .map(|(chain, _)| chain),
    )
}

fn load_calculation_chain_with_conformance_for_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<(CalculationChain, CalculationChainConformance)>> {
    let Some(relationship) = calculation_chain_relationship(package, workbook_uri)? else {
        validate_calculation_chain_part_set(package, None)?;
        return Ok(None);
    };
    validate_calculation_chain_part_set(package, Some(&relationship.part_name))?;
    validate_calculation_chain_part(package, &relationship.part_name)?;
    let part = package.get_part(&relationship.part_name)?;
    let xml = crate::common::mce::process_part(part)?;
    Ok(Some((
        parse_calculation_chain(xml.as_ref())?,
        relationship.conformance,
    )))
}

#[derive(Debug, Clone)]
struct CalculationChainRelationship {
    relationship_id: String,
    part_name: PackURI,
    target_reference: String,
    conformance: CalculationChainConformance,
}

fn calculation_chain_relationship(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<CalculationChainRelationship>> {
    let workbook = package.get_part(workbook_uri)?;
    let mut relationships = workbook.rels().iter().filter(|relationship| {
        matches!(relationship.reltype(), RELATIONSHIP | STRICT_RELATIONSHIP)
    });
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(invalid(
            "workbook has multiple calculation-chain relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("calculation-chain relationship cannot be external"));
    }
    let conformance = if relationship.reltype() == RELATIONSHIP {
        CalculationChainConformance::Transitional
    } else {
        CalculationChainConformance::Strict
    };
    Ok(Some(CalculationChainRelationship {
        relationship_id: relationship.r_id().to_string(),
        part_name: relationship.target_partname()?,
        target_reference: relationship.target_ref().to_string(),
        conformance,
    }))
}

fn validate_calculation_chain_part(package: &OpcPackage, part_name: &PackURI) -> Result<()> {
    let part = package.get_part(part_name)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "calculation-chain part '{part_name}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid("calculation-chain part cannot have relationships"));
    }
    Ok(())
}

fn validate_calculation_chain_part_set(
    package: &OpcPackage,
    relationship_target: Option<&PackURI>,
) -> Result<()> {
    let part_names = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE)
        .map(|part| part.partname().clone())
        .collect::<Vec<_>>();
    if part_names.len() > 1 {
        return Err(invalid(
            "package contains more than one calculation-chain part",
        ));
    }
    match (relationship_target, part_names.as_slice()) {
        (None, []) => Ok(()),
        (None, _) => Err(invalid(
            "package contains a calculation-chain part without a workbook relationship",
        )),
        (Some(_), []) => Ok(()),
        (Some(target), [part_name]) if part_name == target => Ok(()),
        (Some(_), _) => Err(invalid(
            "workbook calculation-chain relationship does not target the calculation-chain part",
        )),
    }
}

fn main_workbook_uri(package: &OpcPackage) -> Result<PackURI> {
    use litchi_opc::constants::content_type as ct;

    let workbook = package.main_document_part()?;
    if !matches!(
        workbook.content_type(),
        ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_SHEET_MACRO_MAIN
            | ct::SML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(invalid(format!(
            "main document part '{}' is not an XML workbook",
            workbook.partname()
        )));
    }
    Ok(workbook.partname().clone())
}

fn next_calculation_chain_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/calcChain.xml".to_string()
        } else {
            format!("/xl/calcChain{suffix}.xml")
        };
        let candidate = PackURI::new(&name).map_err(OoxmlError::InvalidUri)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free calculation-chain part name"))
}

fn next_calculation_chain_relationship_id(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<String> {
    let relationships = package.get_part(workbook_uri)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdCalcChain{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free calculation-chain relationship ID"))
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name == *target)
    })
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    closed: bool,
) -> Result<()> {
    if closed || !is_name(namespace, element, b"calcChain") {
        return Err(invalid(
            "calculation-chain XML has an invalid or trailing root",
        ));
    }
    Ok(())
}

fn is_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    element.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == TRANSITIONAL_NS.as_bytes() || *value == STRICT_NS.as_bytes())
}

fn parse_root_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    chain: &mut CalculationChain,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let raw = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        if raw == "xmlns" || raw.starts_with("xmlns:") {
            if raw != "xmlns" {
                if chain.namespace_declarations.len() >= MAX_EXTENSION_ATTRIBUTES {
                    return Err(invalid("too many calculation-chain namespace declarations"));
                }
                chain.namespace_declarations.push((raw, value));
            }
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if matches!(
            namespace,
            ResolveResult::Unbound | ResolveResult::Unknown(_)
        ) {
            return Err(invalid(format!("unexpected calcChain attribute '{raw}'")));
        }
        push_extension_attribute(&mut chain.extension_attributes, raw, value)?;
    }
    Ok(())
}

fn parse_cell(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<CalculationCell> {
    let mut reference = None;
    let mut sheet_id = None;
    let mut child_chain = None;
    let mut new_level = None;
    let mut new_thread = None;
    let mut array = None;
    let mut extension_attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let raw = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        if raw == "xmlns" || raw.starts_with("xmlns:") {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) {
            match attribute.key.local_name().as_ref() {
                b"r" => set_once(&mut reference, value, "r")?,
                b"i" => set_once(&mut sheet_id, parse_i32(&value, "i")?, "i")?,
                b"s" => set_once(&mut child_chain, parse_bool(&value, "s")?, "s")?,
                b"l" => set_once(&mut new_level, parse_bool(&value, "l")?, "l")?,
                b"t" => set_once(&mut new_thread, parse_bool(&value, "t")?, "t")?,
                b"a" => set_once(&mut array, parse_bool(&value, "a")?, "a")?,
                _ => {
                    return Err(invalid(format!(
                        "unexpected calculation-cell attribute '{raw}'"
                    )));
                },
            }
        } else if matches!(namespace, ResolveResult::Unknown(_)) {
            return Err(invalid(format!(
                "unbound calculation-cell attribute '{raw}'"
            )));
        } else {
            push_extension_attribute(&mut extension_attributes, raw, value)?;
        }
    }
    let reference = reference.ok_or_else(|| invalid("calculation cell requires r"))?;
    validate_reference(&reference)?;
    Ok(CalculationCell {
        reference,
        sheet_id,
        child_chain,
        new_dependency_level: new_level.unwrap_or(false),
        new_thread: new_thread.unwrap_or(false),
        array_formula: array.unwrap_or(false),
        extension_attributes,
    })
}

fn push_cell(chain: &mut CalculationChain, cell: CalculationCell) -> Result<()> {
    if chain.cells.len() >= MAX_CELLS {
        return Err(invalid("calculation chain has too many cells"));
    }
    chain.cells.push(cell);
    Ok(())
}

fn consume_leaf(reader: &mut NsReader<&[u8]>, local: &[u8]) -> Result<()> {
    loop {
        match reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
        {
            Event::End(element) if element.local_name().as_ref() == local => return Ok(()),
            Event::Text(text)
                if text
                    .decode()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                    .trim()
                    .is_empty() => {},
            Event::Comment(_) => {},
            Event::Start(_) | Event::Empty(_) | Event::CData(_) => {
                return Err(invalid("calculation cell must be empty"));
            },
            Event::Eof => return Err(invalid("unterminated calculation cell")),
            _ => return Err(invalid("invalid calculation-cell content")),
        }
    }
}

fn consume_extension_list(reader: &mut NsReader<&[u8]>) -> Result<usize> {
    let mut depth = 1usize;
    let mut nodes = 0usize;
    while depth != 0 {
        match reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
        {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("extension nesting overflow"))?;
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("extension node count overflow"))?;
                if depth > MAX_EXTENSION_DEPTH || nodes > MAX_CELLS {
                    return Err(invalid("calculation-chain extension is too complex"));
                }
            },
            Event::Empty(_) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| invalid("extension node count overflow"))?;
                if nodes > MAX_CELLS {
                    return Err(invalid("calculation-chain extension has too many nodes"));
                }
            },
            Event::End(_) => depth -= 1,
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "DTD and processing instructions are rejected in extensions",
                ));
            },
            Event::Eof => return Err(invalid("unterminated calculation-chain extLst")),
            _ => {},
        }
    }
    position(reader)
}

fn validate_reference(value: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_REFERENCE_BYTES {
        return Err(invalid("calculation-cell reference has invalid length"));
    }
    Cell::reference_to_coords(value)
        .map(|_| ())
        .map_err(|error| invalid(error.to_string()))
}

fn parse_i32(value: &str, name: &str) -> Result<i32> {
    value.parse::<i32>().map_err(|_| {
        invalid(format!(
            "calculation-cell {name} is outside the signed 32-bit bound"
        ))
    })
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!(
            "invalid calculation-cell {name} boolean '{value}'"
        ))),
    }
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(invalid(format!(
            "duplicate calculation-cell {name} attribute"
        )));
    }
    Ok(())
}

fn push_extension_attribute(
    attributes: &mut Vec<CalculationChainExtensionAttribute>,
    qualified_name: String,
    value: String,
) -> Result<()> {
    if attributes.len() >= MAX_EXTENSION_ATTRIBUTES {
        return Err(invalid("too many preserved calculation-chain attributes"));
    }
    if attributes
        .iter()
        .any(|attribute| attribute.qualified_name == qualified_name)
    {
        return Err(invalid(format!(
            "duplicate preserved attribute '{qualified_name}'"
        )));
    }
    attributes.push(CalculationChainExtensionAttribute {
        qualified_name,
        value,
    });
    Ok(())
}

fn write_extension_attributes(
    xml: &mut String,
    attributes: &[CalculationChainExtensionAttribute],
) -> Result<()> {
    if attributes.len() > MAX_EXTENSION_ATTRIBUTES {
        return Err(invalid("too many preserved calculation-chain attributes"));
    }
    for attribute in attributes {
        xml.push(' ');
        xml.push_str(&attribute.qualified_name);
        xml.push_str("=\"");
        escape_attribute(xml, &attribute.value);
        xml.push('"');
    }
    Ok(())
}

fn write_bool_attribute(xml: &mut String, name: &str, value: bool) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str(if value { "=\"1\"" } else { "=\"0\"" });
}

fn escape_attribute(xml: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '"' => xml.push_str("&quot;"),
            '\t' => xml.push_str("&#x9;"),
            '\n' => xml.push_str("&#xA;"),
            '\r' => xml.push_str("&#xD;"),
            _ => xml.push(character),
        }
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("calculation-chain XML offset overflow"))
}

fn raw_range(bytes: &[u8], start: usize, end: usize) -> Result<String> {
    if end < start || end - start > MAX_EXTENSION_BYTES {
        return Err(invalid("calculation-chain extension list is too large"));
    }
    std::str::from_utf8(
        bytes
            .get(start..end)
            .ok_or_else(|| invalid("invalid calculation-chain extension range"))?,
    )
    .map(str::to_owned)
    .map_err(|error| invalid(format!("calculation-chain extension is not UTF-8: {error}")))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::part::{BlobPart, Part};

    #[test]
    fn parses_writes_defaults_inheritance_strict_and_extensions() {
        let xml = br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="urn:test" x:root="v"><c r="A1" i="2" s="1" l="1" t="true" a="1" x:cell="kept"/><c r="B2" s="0"/><extLst><ext uri="urn:test"><x:data value="inert"/></ext></extLst></calcChain>"#;
        let chain = parse_calculation_chain(xml).unwrap();
        assert_eq!(chain.cells().len(), 2);
        assert_eq!(chain.effective_sheet_id(1), Some(2));
        assert_eq!(chain.effective_child_chain(0), Some(true));
        assert_eq!(chain.effective_child_chain(1), Some(false));
        assert!(chain.cells()[0].starts_new_dependency_level());
        assert!(chain.cells()[0].starts_new_thread());
        assert!(chain.cells()[0].is_array_formula());
        let strict = chain.to_xml(CalculationChainConformance::Strict).unwrap();
        assert!(strict.contains(STRICT_NS));
        assert!(strict.contains("x:cell=\"kept\""));
        assert!(strict.contains("<extLst>"));
        let reparsed = parse_calculation_chain(strict.as_bytes()).unwrap();
        assert_eq!(reparsed.cells(), chain.cells());
        assert_eq!(
            reparsed
                .to_xml(CalculationChainConformance::Strict)
                .unwrap(),
            strict
        );
    }

    #[test]
    fn preprocesses_mce_and_rejects_malformed_records() {
        let mce = br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:c/></mc:Choice><mc:Fallback><c r="C3"/></mc:Fallback></mc:AlternateContent></calcChain>"#;
        assert_eq!(
            parse_calculation_chain(mce).unwrap().cells()[0].reference(),
            "C3"
        );
        let invalid = [
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"/>"#),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c/></calcChain>"#),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="XFE1"/></calcChain>"#),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" l="yes"/></calcChain>"#),
            format!(
                r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="2147483648"/></calcChain>"#
            ),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><extLst/><c r="A1"/></calcChain>"#),
            format!(
                r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1"><c r="B1"/></c></calcChain>"#
            ),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" bogus="1"/></calcChain>"#),
        ];
        for xml in invalid {
            assert!(
                parse_calculation_chain(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn stores_rewrites_and_removes_inert_calculation_chain_parts() {
        let mut package = workbook_package();
        let mut chain = CalculationChain::new();
        chain
            .push(
                CalculationCell::new("B2")
                    .unwrap()
                    .set_sheet_id(Some(1))
                    .set_starts_new_dependency_level(true)
                    .clone(),
            )
            .unwrap();

        store_calculation_chain(
            &mut package,
            &chain,
            CalculationChainConformance::Transitional,
        )
        .unwrap();
        assert_eq!(
            load_calculation_chain_from_package(&package).unwrap(),
            Some(chain.clone())
        );

        let workbook = package.main_document_part().unwrap();
        let relationship = workbook
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == RELATIONSHIP)
            .unwrap();
        let relationship_id = relationship.r_id().to_string();
        let part_name = relationship.target_partname().unwrap();
        assert_eq!(part_name, PackURI::new("/xl/calcChain.xml").unwrap());
        assert!(
            std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
                .unwrap()
                .contains(TRANSITIONAL_NS)
        );

        let mut replacement = CalculationChain::new();
        replacement
            .push(CalculationCell::new("C3").unwrap())
            .unwrap();
        store_calculation_chain(
            &mut package,
            &replacement,
            CalculationChainConformance::Strict,
        )
        .unwrap();
        let workbook = package.main_document_part().unwrap();
        let relationship = workbook
            .rels()
            .iter()
            .find(|relationship| relationship.r_id() == relationship_id)
            .unwrap();
        assert_eq!(relationship.reltype(), STRICT_RELATIONSHIP);
        assert_eq!(relationship.target_partname().unwrap(), part_name);
        assert!(
            std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
                .unwrap()
                .contains(STRICT_NS)
        );
        assert_eq!(
            load_calculation_chain_from_package(&package).unwrap(),
            Some(replacement)
        );

        assert!(remove_calculation_chain(&mut package).unwrap());
        assert!(package.get_part(&part_name).is_err());
        assert_eq!(load_calculation_chain_from_package(&package).unwrap(), None);
        assert!(!remove_calculation_chain(&mut package).unwrap());
    }

    #[test]
    fn removal_retains_a_calculation_chain_part_referenced_elsewhere() {
        let mut package = workbook_package();
        let mut chain = CalculationChain::new();
        chain.push(CalculationCell::new("F6").unwrap()).unwrap();
        store_calculation_chain(
            &mut package,
            &chain,
            CalculationChainConformance::Transitional,
        )
        .unwrap();

        let part_name = PackURI::new("/xl/calcChain.xml").unwrap();
        let mut referring_part = BlobPart::new(
            PackURI::new("/xl/retained-reference.xml").unwrap(),
            ct::XML.into(),
            b"<reference/>".to_vec(),
        );
        referring_part.relate_to("calcChain.xml", "urn:litchi:test:calc-chain-reference");
        package.add_part(Box::new(referring_part));

        assert!(remove_calculation_chain(&mut package).unwrap());
        assert!(package.get_part(&part_name).is_ok());
        assert!(load_calculation_chain_from_package(&package).is_err());
        assert!(
            store_calculation_chain(
                &mut package,
                &chain,
                CalculationChainConformance::Transitional,
            )
            .is_err()
        );
    }

    #[test]
    fn workbook_calculation_chain_mutators_refresh_cached_metadata() {
        let mut workbook = crate::xlsx::Workbook::new(workbook_package()).unwrap();
        let mut chain = CalculationChain::new();
        chain.push(CalculationCell::new("D4").unwrap()).unwrap();

        workbook
            .set_calculation_chain(chain.clone(), CalculationChainConformance::Strict)
            .unwrap();
        assert_eq!(workbook.calculation_chain(), Some(&chain));
        assert_eq!(
            workbook.calculation_chain_conformance(),
            Some(CalculationChainConformance::Strict)
        );
        assert_eq!(
            load_calculation_chain_from_package(workbook.opc_package()).unwrap(),
            Some(chain)
        );

        assert!(workbook.remove_calculation_chain().unwrap());
        assert_eq!(workbook.calculation_chain(), None);
        assert!(!workbook.remove_calculation_chain().unwrap());
    }

    #[test]
    fn workbook_calculation_chain_round_trips_through_xlsx_save() {
        let mut workbook = crate::xlsx::Workbook::create().unwrap();
        let mut chain = CalculationChain::new();
        chain
            .push(
                CalculationCell::new("A1")
                    .unwrap()
                    .set_sheet_id(Some(1))
                    .clone(),
            )
            .unwrap();
        workbook
            .set_calculation_chain(chain.clone(), CalculationChainConformance::Transitional)
            .unwrap();

        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("calculation-chain.xlsx");
        workbook.save(&path).unwrap();
        let reopened = crate::xlsx::Workbook::open(&path).unwrap();
        assert_eq!(reopened.calculation_chain(), Some(&chain));
        assert_eq!(
            reopened.calculation_chain_conformance(),
            Some(CalculationChainConformance::Transitional)
        );
        assert_eq!(
            reopened
                .opc_package()
                .iter_parts()
                .filter(|part| part.content_type() == CONTENT_TYPE)
                .count(),
            1
        );
    }

    #[test]
    fn package_calculation_chain_mutators_reject_invalid_existing_graphs() {
        let mut package = synthetic_package(RELATIONSHIP, false, ct::XML, false);
        let chain_part = PackURI::new("/xl/calcChain.xml").unwrap();
        let original = package.get_part(&chain_part).unwrap().blob().to_vec();
        let mut chain = CalculationChain::new();
        chain.push(CalculationCell::new("E5").unwrap()).unwrap();

        assert!(
            store_calculation_chain(
                &mut package,
                &chain,
                CalculationChainConformance::Transitional,
            )
            .is_err()
        );
        assert_eq!(package.get_part(&chain_part).unwrap().blob(), original);
        assert!(remove_calculation_chain(&mut package).is_err());
        assert!(package.get_part(&chain_part).is_ok());

        let mut duplicate = synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, false);
        duplicate
            .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                RELATIONSHIP.into(),
                "calcChain.xml".into(),
                "rIdDuplicateCalcChain".into(),
                false,
            );
        assert!(
            store_calculation_chain(
                &mut duplicate,
                &chain,
                CalculationChainConformance::Transitional,
            )
            .is_err()
        );
        assert!(remove_calculation_chain(&mut duplicate).is_err());

        let mut duplicate_part = synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, false);
        duplicate_part.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/calcChainExtra.xml").unwrap(),
            CONTENT_TYPE.into(),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="F6"/></calcChain>"#).into_bytes(),
        )));
        assert!(load_calculation_chain_from_package(&duplicate_part).is_err());
        assert!(
            store_calculation_chain(
                &mut duplicate_part,
                &chain,
                CalculationChainConformance::Transitional,
            )
            .is_err()
        );
        assert!(remove_calculation_chain(&mut duplicate_part).is_err());

        let mut external = synthetic_package(RELATIONSHIP, true, CONTENT_TYPE, false);
        assert!(
            store_calculation_chain(
                &mut external,
                &chain,
                CalculationChainConformance::Transitional,
            )
            .is_err()
        );
        assert!(remove_calculation_chain(&mut external).is_err());
    }

    #[test]
    fn loads_real_poi_and_synthetic_packages_and_validates_relationships() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..//test-data/poi/test-data/spreadsheet/62834.xlsx");
        let workbook = crate::xlsx::Workbook::open(path).unwrap();
        let chain = workbook.calculation_chain().unwrap();
        assert_eq!(chain.cells().len(), 3);
        assert_eq!(chain.cells()[0].reference(), "A5");
        assert!(chain.cells()[0].starts_new_dependency_level());
        assert_eq!(chain.cells()[2].child_chain_override(), Some(true));

        let package = synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, false);
        let workbook = crate::xlsx::Workbook::new(package).unwrap();
        assert_eq!(
            workbook.calculation_chain().unwrap().cells()[0].reference(),
            "A1"
        );

        assert!(
            crate::xlsx::Workbook::new(synthetic_package(RELATIONSHIP, true, CONTENT_TYPE, false))
                .is_err()
        );
        assert!(
            crate::xlsx::Workbook::new(synthetic_package(RELATIONSHIP, false, ct::XML, false))
                .is_err()
        );
        assert!(
            crate::xlsx::Workbook::new(synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, true))
                .is_err()
        );
    }

    fn workbook_package() -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let workbook = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.into(),
            format!(r#"<workbook xmlns="{TRANSITIONAL_NS}"><sheets/></workbook>"#).into_bytes(),
        );
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook));
        package
    }

    fn synthetic_package(
        relationship_type: &str,
        external: bool,
        content_type: &str,
        outbound: bool,
    ) -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let mut workbook = BlobPart::new(
            workbook_uri.clone(),
            ct::SML_SHEET_MAIN.into(),
            format!(r#"<workbook xmlns="{TRANSITIONAL_NS}"><sheets/></workbook>"#).into_bytes(),
        );
        if external {
            workbook.relate_to_ext("https://example.invalid/calcChain.xml", relationship_type);
        } else {
            workbook.relate_to("calcChain.xml", relationship_type);
        }
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook));
        if !external {
            let mut chain = BlobPart::new(
                PackURI::new("/xl/calcChain.xml").unwrap(),
                content_type.into(),
                format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"/></calcChain>"#)
                    .into_bytes(),
            );
            if outbound {
                chain.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
            }
            package.add_part(Box::new(chain));
        }
        package
    }
}
