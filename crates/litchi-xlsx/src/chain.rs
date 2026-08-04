//! Typed, inert SpreadsheetML calculation-chain metadata and package ownership.

use std::collections::HashSet;

use crate::error::{Error, Result, allocation, invalid};
use litchi_opc::{OpcPackage, PackURI};
use litchi_sheet::{At, Cell as Address};
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
const MAX_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;
const MAX_EXTENSION_BYTES: usize = 16 * 1024 * 1024;
const MAX_CELL_CONTENT_BYTES: usize = 16 * 1024 * 1024;
const MAX_EXTENSION_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1024 * 1024;
const MAX_EXTENSION_DEPTH: usize = 128;
const MAX_REFERENCE_BYTES: usize = 32;

/// Namespace family used by the calculation-chain writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Conformance {
    #[default]
    Transitional,
    Strict,
}

impl Conformance {
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

/// One Excel sheet identifier proven to be in the native `1..=65534` domain.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Sheet(u16);

impl Sheet {
    /// Validate a native sheet identifier.
    pub fn new(value: u32) -> Result<Self> {
        let value = u16::try_from(value)
            .ok()
            .filter(|value| (1..=65_534).contains(value))
            .ok_or_else(|| {
                invalid(format!(
                    "calculation-chain sheet ID {value} is outside 1..=65534"
                ))
            })?;
        Ok(Self(value))
    }

    /// Return the native one-based sheet identifier.
    pub const fn get(self) -> u16 {
        self.0
    }
}

impl TryFrom<u32> for Sheet {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

/// Dependency role of one calculation-chain cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[non_exhaustive]
pub enum Step {
    /// Continue the current dependency level as a parent formula.
    #[default]
    Same,
    /// Start a new dependency level.
    Level,
    /// Continue the current level as a child formula.
    Child,
}

bitflags::bitflags! {
    /// Orthogonal calculation-cell markers.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct Flags: u8 {
        /// Retain the deprecated producer thread marker.
        const THREAD = 1 << 0;
        /// The cell belongs to an array formula.
        const ARRAY = 1 << 1;
    }
}

/// Preserved extension vocabulary.
pub mod raw {
    /// An MCE-preserved, non-schema attribute retained without interpretation.
    #[derive(Debug, Clone, PartialEq, Eq)]
    pub struct Attr {
        pub(super) name: String,
        pub(super) value: String,
    }

    impl Attr {
        /// Return the original qualified attribute name.
        pub fn name(&self) -> &str {
            &self.name
        }

        /// Return the decoded attribute value.
        pub fn value(&self) -> &str {
            &self.value
        }
    }
}

use raw::Attr;

/// One formula cell in calculation order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cell {
    reference: Box<str>,
    address: Address,
    sheet: Sheet,
    explicit_sheet: bool,
    step: Step,
    flags: Flags,
    attrs: Vec<Attr>,
}

impl Cell {
    /// Create a cell with an explicit sheet and checked A1 address.
    pub fn new<'a>(sheet: Sheet, at: impl Into<At<'a>>) -> Result<Self> {
        let address = at.into().resolve()?;
        Ok(Self {
            reference: address.a1().into_boxed_str(),
            address,
            sheet,
            explicit_sheet: true,
            step: Step::Same,
            flags: Flags::empty(),
            attrs: Vec::new(),
        })
    }

    /// Return the original checked A1 spelling.
    pub fn reference(&self) -> &str {
        &self.reference
    }

    /// Return the typed grid address.
    pub const fn address(&self) -> Address {
        self.address
    }

    /// Return the effective sheet identifier.
    pub const fn sheet(&self) -> Sheet {
        self.sheet
    }

    /// Return this cell's mutually exclusive dependency role.
    pub const fn step(&self) -> Step {
        self.step
    }

    /// Return orthogonal producer markers.
    pub const fn flags(&self) -> Flags {
        self.flags
    }

    /// Set the mutually exclusive dependency role.
    pub fn set_step(&mut self, step: Step) -> &mut Self {
        self.step = step;
        self
    }

    /// Set the orthogonal producer markers.
    pub fn set_flags(&mut self, flags: Flags) -> &mut Self {
        self.flags = flags;
        self
    }

    /// Return bounded preserved attributes.
    pub fn attrs(&self) -> &[Attr] {
        &self.attrs
    }
}

/// Ordered metadata from the workbook's single Calculation Chain part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Chain {
    cells: Vec<Cell>,
    ambiguous_key: Option<(Sheet, Address)>,
    extension_list_xml: Option<String>,
    namespace_declarations: Vec<(String, String)>,
    attrs: Vec<Attr>,
}

impl Chain {
    /// Create a non-empty chain. The first cell is always written with a sheet ID.
    pub fn new(mut first: Cell) -> Self {
        first.explicit_sheet = true;
        Self {
            cells: vec![first],
            ambiguous_key: None,
            extension_list_xml: None,
            namespace_declarations: Vec::new(),
            attrs: Vec::new(),
        }
    }

    /// Borrow cells in calculation order.
    pub fn cells(&self) -> &[Cell] {
        &self.cells
    }

    /// Return the number of calculation cells.
    pub fn len(&self) -> usize {
        self.cells.len()
    }

    /// A chain is statically non-empty.
    pub const fn is_empty(&self) -> bool {
        false
    }

    /// Look up a semantic sheet/address key. Duplicate malformed input is an error.
    pub fn get<'a>(&self, sheet: Sheet, at: impl Into<At<'a>>) -> Result<Option<&Cell>> {
        let address = at.into().resolve()?;
        Ok(self
            .matching_position(sheet, address)?
            .and_then(|position| self.cells.get(position)))
    }

    /// Borrow a checked calculation-order position.
    pub fn at(&self, position: usize) -> Result<&Cell> {
        self.cells.get(position).ok_or_else(|| {
            invalid(format!(
                "calculation-chain position {position} is outside 0..{}",
                self.cells.len()
            ))
        })
    }

    /// Append a unique semantic cell.
    pub fn push(&mut self, cell: Cell) -> Result<&mut Self> {
        if self.cells.len() >= MAX_CELLS {
            return Err(invalid("calculation chain has too many cells"));
        }
        if self.matching_position(cell.sheet, cell.address)?.is_some() {
            return Err(invalid(format!(
                "calculation cell {} on sheet {} already exists",
                cell.address,
                cell.sheet.get()
            )));
        }
        self.cells
            .try_reserve(1)
            .map_err(|source| allocation("calculation-chain cells", source))?;
        self.cells.push(cell);
        self.ensure_sheet_boundaries();
        Ok(self)
    }

    /// Insert a unique cell at a checked calculation-order position.
    pub fn insert(&mut self, position: usize, cell: Cell) -> Result<&mut Self> {
        if position > self.cells.len() {
            return Err(invalid(format!(
                "calculation-chain insertion position {position} is outside 0..={}",
                self.cells.len()
            )));
        }
        if self.cells.len() >= MAX_CELLS {
            return Err(invalid("calculation chain has too many cells"));
        }
        if self.matching_position(cell.sheet, cell.address)?.is_some() {
            return Err(invalid(format!(
                "calculation cell {} on sheet {} already exists",
                cell.address,
                cell.sheet.get()
            )));
        }
        self.cells
            .try_reserve(1)
            .map_err(|source| allocation("calculation-chain cells", source))?;
        self.cells.insert(position, cell);
        self.ensure_sheet_boundaries();
        Ok(self)
    }

    /// Insert or replace by semantic sheet/address key, preserving existing order.
    pub fn put(&mut self, cell: Cell) -> Result<Option<Cell>> {
        match self.matching_position(cell.sheet, cell.address)? {
            None => {
                if self.cells.len() >= MAX_CELLS {
                    return Err(invalid("calculation chain has too many cells"));
                }
                self.cells
                    .try_reserve(1)
                    .map_err(|source| allocation("calculation-chain cells", source))?;
                self.cells.push(cell);
                self.ensure_sheet_boundaries();
                Ok(None)
            },
            Some(position) => {
                let previous = std::mem::replace(&mut self.cells[position], cell);
                self.ensure_sheet_boundaries();
                Ok(Some(previous))
            },
        }
    }

    /// Replace a checked calculation-order position.
    pub fn replace_at(&mut self, position: usize, cell: Cell) -> Result<Cell> {
        self.at(position)?;
        self.reject_duplicate(&cell, Some(position))?;
        let mut key_index = self.key_index_if_ambiguous()?;
        let previous = std::mem::replace(&mut self.cells[position], cell);
        if let Some(key_index) = &mut key_index {
            self.refresh_ambiguity(key_index);
        }
        self.ensure_sheet_boundaries();
        Ok(previous)
    }

    /// Remove a semantic sheet/address key, while retaining a non-empty chain.
    pub fn remove<'a>(&mut self, sheet: Sheet, at: impl Into<At<'a>>) -> Result<Option<Cell>> {
        let address = at.into().resolve()?;
        match self.matching_position(sheet, address)? {
            None => Ok(None),
            Some(position) => self.remove_at(position).map(Some),
        }
    }

    /// Remove a checked calculation-order position.
    pub fn remove_at(&mut self, position: usize) -> Result<Cell> {
        self.at(position)?;
        if self.cells.len() == 1 {
            return Err(invalid("a calculation chain cannot be empty"));
        }
        let mut key_index = self.key_index_if_ambiguous()?;
        let removed = self.cells.remove(position);
        if let Some(key_index) = &mut key_index {
            self.refresh_ambiguity(key_index);
        }
        self.ensure_sheet_boundaries();
        Ok(removed)
    }

    /// Move one checked position, interpreting `to` in the final sequence.
    pub fn move_at(&mut self, from: usize, to: usize) -> Result<&mut Self> {
        self.at(from)?;
        if to >= self.cells.len() {
            return Err(invalid(format!(
                "calculation-chain destination {to} is outside 0..{}",
                self.cells.len()
            )));
        }
        if from != to {
            let cell = self.cells.remove(from);
            self.cells.insert(to, cell);
            self.ensure_sheet_boundaries();
        }
        Ok(self)
    }

    /// Return the bounded, preserved extension-list XML, when present.
    pub fn extension_list_xml(&self) -> Option<&str> {
        self.extension_list_xml.as_deref()
    }

    /// Return bounded, preserved attributes from the chain root.
    pub fn attrs(&self) -> &[Attr] {
        &self.attrs
    }

    fn matching_position(&self, sheet: Sheet, address: Address) -> Result<Option<usize>> {
        if let Some((ambiguous_sheet, ambiguous_address)) = self.ambiguous_key {
            return Err(invalid(format!(
                "calculation chain contains ambiguous cell {} on sheet {}",
                ambiguous_address,
                ambiguous_sheet.get()
            )));
        }
        Ok(self
            .cells
            .iter()
            .position(|cell| cell.sheet == sheet && cell.address == address))
    }

    fn key_index_if_ambiguous(&self) -> Result<Option<HashSet<(Sheet, Address)>>> {
        if self.ambiguous_key.is_none() {
            return Ok(None);
        }
        let mut seen = HashSet::new();
        seen.try_reserve(self.cells.len())
            .map_err(|source| allocation("calculation-chain key index", source))?;
        Ok(Some(seen))
    }

    fn refresh_ambiguity(&mut self, seen: &mut HashSet<(Sheet, Address)>) {
        seen.clear();
        self.ambiguous_key = None;
        for cell in &self.cells {
            let key = (cell.sheet, cell.address);
            if !seen.insert(key) {
                self.ambiguous_key = Some(key);
                break;
            }
        }
    }

    fn reject_duplicate(&self, cell: &Cell, except: Option<usize>) -> Result<()> {
        if self.cells.iter().enumerate().any(|(position, existing)| {
            Some(position) != except
                && existing.sheet == cell.sheet
                && existing.address == cell.address
        }) {
            return Err(invalid(format!(
                "calculation cell {} on sheet {} already exists",
                cell.address,
                cell.sheet.get()
            )));
        }
        Ok(())
    }

    fn ensure_sheet_boundaries(&mut self) {
        if let Some(first) = self.cells.first_mut() {
            first.explicit_sheet = true;
        }
        for position in 1..self.cells.len() {
            if self.cells[position - 1].sheet != self.cells[position].sheet {
                self.cells[position].explicit_sheet = true;
            }
        }
    }
}

/// Serialize a complete calculation chain with bounded allocation.
pub fn write(chain: &Chain, conformance: Conformance) -> Result<Vec<u8>> {
    if chain.cells.len() > MAX_CELLS {
        return Err(invalid("calculation chain has too many cells"));
    }
    let capacity = wire_len(chain, conformance)?;
    let mut xml = String::new();
    xml.try_reserve_exact(capacity)
        .map_err(|source| allocation("calculation-chain output", source))?;
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str("<calcChain xmlns=\"");
    xml.push_str(conformance.namespace());
    xml.push('"');
    for (name, value) in &chain.namespace_declarations {
        if name != "xmlns" {
            xml.push(' ');
            xml.push_str(name);
            xml.push_str("=\"");
            escape_attribute(&mut xml, value);
            xml.push('"');
        }
    }
    write_extension_attributes(&mut xml, &chain.attrs)?;
    xml.push('>');
    for cell in &chain.cells {
        xml.push_str("<c r=\"");
        escape_attribute(&mut xml, &cell.reference);
        xml.push('"');
        if cell.explicit_sheet {
            xml.push_str(" i=\"");
            push_u16(&mut xml, cell.sheet.get());
            xml.push('"');
        }
        match cell.step {
            Step::Same => {},
            Step::Level => write_bool_attribute(&mut xml, "l", true),
            Step::Child => write_bool_attribute(&mut xml, "s", true),
        }
        if cell.flags.contains(Flags::THREAD) {
            write_bool_attribute(&mut xml, "t", true);
        }
        if cell.flags.contains(Flags::ARRAY) {
            write_bool_attribute(&mut xml, "a", true);
        }
        write_extension_attributes(&mut xml, &cell.attrs)?;
        xml.push_str("/>");
    }
    if let Some(extension) = &chain.extension_list_xml {
        xml.push_str(extension);
    }
    xml.push_str("</calcChain>");
    Ok(xml.into_bytes())
}

#[derive(Default)]
struct Builder {
    cells: Vec<Cell>,
    seen_keys: HashSet<(Sheet, Address)>,
    ambiguous_key: Option<(Sheet, Address)>,
    extension_list_xml: Option<String>,
    namespace_declarations: Vec<(String, String)>,
    attrs: Vec<Attr>,
}

impl Builder {
    fn finish(self) -> Result<Chain> {
        if self.cells.is_empty() {
            return Err(invalid("calculation chain must contain at least one cell"));
        }
        let mut chain = Chain {
            cells: self.cells,
            ambiguous_key: self.ambiguous_key,
            extension_list_xml: self.extension_list_xml,
            namespace_declarations: self.namespace_declarations,
            attrs: self.attrs,
        };
        chain.ensure_sheet_boundaries();
        Ok(chain)
    }
}

/// Parse an isolated Calculation Chain part. Formula text is never evaluated.
pub fn read(xml: &[u8]) -> Result<Chain> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid(format!(
            "calculation-chain XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)
        .map_err(|error| invalid(format!("calculation-chain MCE error: {error}")))?;
    let bytes = processed.as_ref();
    if bytes.len() > MAX_XML_BYTES {
        return Err(invalid(format!(
            "processed calculation-chain XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut builder = Builder::default();
    let mut current_sheet = None;
    let mut saw_root = false;
    let mut closed_root = false;
    let mut saw_extensions = false;
    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid calculation-chain XML: {error}")))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if !saw_root => {
                validate_root(&namespace, &element, closed_root)?;
                saw_root = true;
                parse_root_attributes(&element, decoder, &resolver, &mut builder)?;
            },
            Event::Empty(element) if !saw_root => {
                validate_root(&namespace, &element, closed_root)?;
                saw_root = true;
                closed_root = true;
                parse_root_attributes(&element, decoder, &resolver, &mut builder)?;
            },
            Event::Empty(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"c") =>
            {
                if saw_extensions {
                    return Err(invalid("calculation cells must precede extLst"));
                }
                let cell = parse_cell(&element, decoder, &resolver, current_sheet)?;
                current_sheet = Some(cell.sheet);
                push_cell(&mut builder, cell)?;
            },
            Event::Start(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"c") =>
            {
                if saw_extensions {
                    return Err(invalid("calculation cells must precede extLst"));
                }
                let cell = parse_cell(&element, decoder, &resolver, current_sheet)?;
                let content_start = position(&reader)?;
                consume_leaf(&mut reader, b"c", content_start)?;
                current_sheet = Some(cell.sheet);
                push_cell(&mut builder, cell)?;
            },
            Event::Empty(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"extLst") =>
            {
                if std::mem::replace(&mut saw_extensions, true) {
                    return Err(invalid("duplicate calculation-chain extLst"));
                }
                let end = position(&reader)?;
                builder.extension_list_xml = Some(raw_range(bytes, start, end)?);
            },
            Event::Start(element)
                if saw_root && !closed_root && is_name(&namespace, &element, b"extLst") =>
            {
                if std::mem::replace(&mut saw_extensions, true) {
                    return Err(invalid("duplicate calculation-chain extLst"));
                }
                let end = consume_extension_list(&mut reader, start)?;
                builder.extension_list_xml = Some(raw_range(bytes, start, end)?);
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
                    .map_err(|error| invalid(format!("invalid calculation-chain text: {error}")))?
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
    builder.finish()
}

/// Load the optional inert calculation chain and its relationship conformance.
/// Formula cells are parsed as metadata only; no formula is evaluated.
pub fn load(package: &OpcPackage) -> Result<Option<(Chain, Conformance)>> {
    let workbook_uri = main_workbook_uri(package)?;
    load_for_workbook(package, &workbook_uri)
}

/// Validate the package topology for the optional calculation chain without
/// decoding its inert XML payload.
pub(crate) fn validate_package(package: &OpcPackage) -> Result<()> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(relationship) = relationship(package, &workbook_uri)? else {
        validate_part_set(package, None)?;
        return Ok(());
    };
    validate_part_set(package, Some(&relationship.part_name))?;
    validate_part(package, &relationship.part_name)
}

/// Store a caller-authored inert calculation chain in a SpreadsheetML package.
///
/// The supplied order is serialized without recalculating formulas or inferring
/// dependencies. Existing calculation-chain graph violations are rejected
/// before any package part is changed. The requested conformance is applied to
/// both the part XML and its workbook relationship.
pub fn put(package: &mut OpcPackage, chain: &Chain, conformance: Conformance) -> Result<bool> {
    let xml = write(chain, conformance)?;
    let workbook_uri = main_workbook_uri(package)?;
    let existing = relationship(package, &workbook_uri)?;

    if let Some(existing) = existing {
        validate_part_set(package, Some(&existing.part_name))?;
        validate_part(package, &existing.part_name)?;
        let bytes_changed = package.get_part(&existing.part_name)?.blob() != xml;
        let relationship_changed = existing.conformance != conformance;
        if !bytes_changed && !relationship_changed {
            return Ok(false);
        }
        if bytes_changed {
            package.get_part_mut(&existing.part_name)?.set_blob(xml);
        }
        if relationship_changed {
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
        validate_part_set(package, None)?;
        let part_name = next_part_name(package)?;
        let relationship_id = next_relationship_id(package, &workbook_uri)?;
        let target = part_name.relative_ref(workbook_uri.base_uri());
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name.clone(),
            CONTENT_TYPE.into(),
            xml,
        )))?;
        let workbook = match package.get_part_mut(&workbook_uri) {
            Ok(workbook) => workbook,
            Err(error) => {
                package.remove_part(&part_name);
                return Err(error.into());
            },
        };
        workbook.rels_mut().add_relationship(
            conformance.relationship_type().into(),
            target,
            relationship_id,
            false,
        );
    }

    package.unsign();
    Ok(true)
}

/// Remove the workbook's calculation-chain relationship and its unreferenced part.
///
/// No formulas are changed. A target that is also referenced elsewhere in the
/// package is retained.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(existing) = relationship(package, &workbook_uri)? else {
        validate_part_set(package, None)?;
        return Ok(false);
    };
    validate_part_set(package, Some(&existing.part_name))?;
    validate_part(package, &existing.part_name)?;
    let retain_part = part_is_referenced_elsewhere(
        package,
        &existing.part_name,
        &workbook_uri,
        &existing.relationship_id,
    )?;

    package
        .get_part_mut(&workbook_uri)?
        .rels_mut()
        .remove(&existing.relationship_id);
    if !retain_part {
        package.remove_part(&existing.part_name);
    }
    package.unsign();
    Ok(true)
}

fn load_for_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<(Chain, Conformance)>> {
    let Some(relationship) = relationship(package, workbook_uri)? else {
        validate_part_set(package, None)?;
        return Ok(None);
    };
    validate_part_set(package, Some(&relationship.part_name))?;
    validate_part(package, &relationship.part_name)?;
    let part = package.get_part(&relationship.part_name)?;
    Ok(Some((read(part.blob())?, relationship.conformance)))
}

#[derive(Debug, Clone)]
struct Relationship {
    relationship_id: String,
    part_name: PackURI,
    target_reference: String,
    conformance: Conformance,
}

fn relationship(package: &OpcPackage, workbook_uri: &PackURI) -> Result<Option<Relationship>> {
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
        Conformance::Transitional
    } else {
        Conformance::Strict
    };
    Ok(Some(Relationship {
        relationship_id: relationship.r_id().to_string(),
        part_name: relationship.target_partname()?,
        target_reference: relationship.target_ref().to_string(),
        conformance,
    }))
}

fn validate_part(package: &OpcPackage, part_name: &PackURI) -> Result<()> {
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

fn validate_part_set(package: &OpcPackage, relationship_target: Option<&PackURI>) -> Result<()> {
    let mut parts = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE);
    let part_name = parts.next().map(|part| part.partname());
    if parts.next().is_some() {
        return Err(invalid(
            "package contains more than one calculation-chain part",
        ));
    }
    match (relationship_target, part_name) {
        (None, None) => Ok(()),
        (None, Some(_)) => Err(invalid(
            "package contains a calculation-chain part without a workbook relationship",
        )),
        (Some(_), None) => Ok(()),
        (Some(target), Some(part_name)) if part_name == target => Ok(()),
        (Some(_), Some(_)) => Err(invalid(
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

fn next_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/calcChain.xml".to_string()
        } else {
            format!("/xl/calcChain{suffix}.xml")
        };
        let candidate = PackURI::new(&name).map_err(invalid)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free calculation-chain part name"))
}

fn next_relationship_id(package: &OpcPackage, workbook_uri: &PackURI) -> Result<String> {
    let relationships = package.get_part(workbook_uri)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdCalcChain{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free calculation-chain relationship ID"))
}

fn part_is_referenced_elsewhere(
    package: &OpcPackage,
    target: &PackURI,
    owner: &PackURI,
    owner_relationship: &str,
) -> Result<bool> {
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if part.partname() == owner && relationship.r_id() == owner_relationship {
                continue;
            }
            if !relationship.is_external() && relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    Ok(false)
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
    builder: &mut Builder,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid calcChain attribute: {error}")))?;
        let raw = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("calcChain attribute name is not UTF-8: {error}")))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(format!("invalid calcChain attribute value: {error}")))?
            .into_owned();
        validate_attribute_size(&raw, &value)?;
        if raw == "xmlns" || raw.starts_with("xmlns:") {
            if raw != "xmlns" {
                if builder.namespace_declarations.len() >= MAX_EXTENSION_ATTRIBUTES {
                    return Err(invalid("too many calculation-chain namespace declarations"));
                }
                builder.namespace_declarations.push((raw, value));
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
        push_extension_attribute(&mut builder.attrs, raw, value)?;
    }
    Ok(())
}

fn parse_cell(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    inherited_sheet: Option<Sheet>,
) -> Result<Cell> {
    let mut reference = None;
    let mut sheet = None;
    let mut child = None;
    let mut new_level = None;
    let mut thread = None;
    let mut array = None;
    let mut attrs = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid calculation-cell attribute: {error}")))?;
        let raw = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| {
                invalid(format!(
                    "calculation-cell attribute name is not UTF-8: {error}"
                ))
            })?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(format!("invalid calculation-cell value: {error}")))?
            .into_owned();
        validate_attribute_size(&raw, &value)?;
        if raw == "xmlns" || raw.starts_with("xmlns:") {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) {
            match attribute.key.local_name().as_ref() {
                b"r" => set_once(&mut reference, value, "r")?,
                b"i" => set_once(&mut sheet, parse_sheet(&value)?, "i")?,
                b"s" => set_once(&mut child, parse_bool(&value, "s")?, "s")?,
                b"l" => set_once(&mut new_level, parse_bool(&value, "l")?, "l")?,
                b"t" => set_once(&mut thread, parse_bool(&value, "t")?, "t")?,
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
            push_extension_attribute(&mut attrs, raw, value)?;
        }
    }
    let reference = reference.ok_or_else(|| invalid("calculation cell requires r"))?;
    let address = parse_reference(&reference)?;
    let explicit_sheet = sheet.is_some();
    let sheet = sheet.or(inherited_sheet).ok_or_else(|| {
        invalid("the first calculation-chain cell must specify sheet attribute i")
    })?;
    let child = child.unwrap_or(false);
    let new_level = new_level.unwrap_or(false);
    if child && new_level {
        return Err(invalid(
            "calculation-cell attributes l and s are mutually exclusive",
        ));
    }
    let step = if child {
        Step::Child
    } else if new_level {
        Step::Level
    } else {
        Step::Same
    };
    let mut flags = Flags::empty();
    flags.set(Flags::THREAD, thread.unwrap_or(false));
    flags.set(Flags::ARRAY, array.unwrap_or(false));
    Ok(Cell {
        reference: reference.into_boxed_str(),
        address,
        sheet,
        explicit_sheet,
        step,
        flags,
        attrs,
    })
}

fn push_cell(builder: &mut Builder, cell: Cell) -> Result<()> {
    if builder.cells.len() >= MAX_CELLS {
        return Err(invalid("calculation chain has too many cells"));
    }
    let key = (cell.sheet, cell.address);
    let duplicate = builder.seen_keys.contains(&key);
    builder
        .cells
        .try_reserve(1)
        .map_err(|source| allocation("calculation-chain cells", source))?;
    if !duplicate {
        builder
            .seen_keys
            .try_reserve(1)
            .map_err(|source| allocation("calculation-chain key index", source))?;
        builder.seen_keys.insert(key);
    } else if builder.ambiguous_key.is_none() {
        builder.ambiguous_key = Some(key);
    }
    builder.cells.push(cell);
    Ok(())
}

fn consume_leaf(reader: &mut NsReader<&[u8]>, local: &[u8], start: usize) -> Result<()> {
    loop {
        let event_start = position(reader)?;
        enforce_budget(
            start,
            event_start,
            MAX_CELL_CONTENT_BYTES,
            "calculation cell content exceeds its byte limit",
        )?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid calculation-cell XML: {error}")))?;
        let event_end = position(reader)?;
        enforce_budget(
            start,
            event_end,
            MAX_CELL_CONTENT_BYTES,
            "calculation cell content exceeds its byte limit",
        )?;
        match event {
            Event::End(element) if element.local_name().as_ref() == local => return Ok(()),
            Event::Text(text)
                if text
                    .decode()
                    .map_err(|error| invalid(format!("invalid calculation-cell text: {error}")))?
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

fn consume_extension_list(reader: &mut NsReader<&[u8]>, start: usize) -> Result<usize> {
    let mut depth = 1usize;
    let mut nodes = 0usize;
    while depth != 0 {
        let event_start = position(reader)?;
        enforce_budget(
            start,
            event_start,
            MAX_EXTENSION_BYTES,
            "calculation-chain extension list is too large",
        )?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid extension XML: {error}")))?;
        let event_end = position(reader)?;
        enforce_budget(
            start,
            event_end,
            MAX_EXTENSION_BYTES,
            "calculation-chain extension list is too large",
        )?;
        match event {
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

fn parse_reference(value: &str) -> Result<Address> {
    if value.is_empty() || value.len() > MAX_REFERENCE_BYTES {
        return Err(invalid("calculation-cell reference has invalid length"));
    }
    Address::from_a1(value).map_err(Into::into)
}

fn parse_sheet(value: &str) -> Result<Sheet> {
    let value = value
        .parse::<u32>()
        .map_err(|_| invalid("calculation-cell i is not an unsigned integer"))?;
    Sheet::new(value)
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

fn push_extension_attribute(attributes: &mut Vec<Attr>, name: String, value: String) -> Result<()> {
    validate_attribute_size(&name, &value)?;
    if attributes.len() >= MAX_EXTENSION_ATTRIBUTES {
        return Err(invalid("too many preserved calculation-chain attributes"));
    }
    if attributes.iter().any(|attribute| attribute.name == name) {
        return Err(invalid(format!("duplicate preserved attribute '{name}'")));
    }
    attributes.push(Attr { name, value });
    Ok(())
}

fn write_extension_attributes(xml: &mut String, attributes: &[Attr]) -> Result<()> {
    if attributes.len() > MAX_EXTENSION_ATTRIBUTES {
        return Err(invalid("too many preserved calculation-chain attributes"));
    }
    for attribute in attributes {
        xml.push(' ');
        xml.push_str(&attribute.name);
        xml.push_str("=\"");
        escape_attribute(xml, &attribute.value);
        xml.push('"');
    }
    Ok(())
}

fn validate_attribute_size(name: &str, value: &str) -> Result<()> {
    if name.len() > MAX_ATTRIBUTE_BYTES || value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(invalid(format!(
            "calculation-chain attribute exceeds {MAX_ATTRIBUTE_BYTES} bytes"
        )));
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

fn wire_len(chain: &Chain, conformance: Conformance) -> Result<usize> {
    let mut len = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.len();
    add_len(&mut len, "<calcChain xmlns=\"".len())?;
    add_len(&mut len, conformance.namespace().len())?;
    add_len(&mut len, 1)?;
    if chain.namespace_declarations.len() > MAX_EXTENSION_ATTRIBUTES
        || chain.attrs.len() > MAX_EXTENSION_ATTRIBUTES
    {
        return Err(invalid("too many calculation-chain root attributes"));
    }
    for (name, value) in &chain.namespace_declarations {
        validate_attribute_size(name, value)?;
        if name != "xmlns" {
            add_len(&mut len, attribute_len(name, value)?)?;
        }
    }
    for attribute in &chain.attrs {
        add_len(&mut len, attribute_len(&attribute.name, &attribute.value)?)?;
    }
    add_len(&mut len, 1)?;
    for (position, cell) in chain.cells.iter().enumerate() {
        if position == 0 && !cell.explicit_sheet {
            return Err(invalid(
                "the first calculation-chain cell must carry an explicit sheet ID",
            ));
        }
        if cell.reference.len() > MAX_REFERENCE_BYTES
            || parse_reference(&cell.reference)? != cell.address
        {
            return Err(invalid(
                "calculation-cell reference no longer matches its address",
            ));
        }
        if cell.attrs.len() > MAX_EXTENSION_ATTRIBUTES {
            return Err(invalid("too many calculation-cell extension attributes"));
        }
        add_len(&mut len, "<c r=\"".len())?;
        add_len(&mut len, escaped_len(&cell.reference)?)?;
        add_len(&mut len, 1)?;
        if cell.explicit_sheet {
            add_len(&mut len, " i=\"".len())?;
            add_len(&mut len, decimal_len(u32::from(cell.sheet.get())))?;
            add_len(&mut len, 1)?;
        }
        if cell.step != Step::Same {
            add_len(&mut len, 6)?;
        }
        add_len(
            &mut len,
            usize::from(cell.flags.contains(Flags::THREAD)) * 6,
        )?;
        add_len(&mut len, usize::from(cell.flags.contains(Flags::ARRAY)) * 6)?;
        for attribute in &cell.attrs {
            add_len(&mut len, attribute_len(&attribute.name, &attribute.value)?)?;
        }
        add_len(&mut len, 2)?;
    }
    if let Some(extension) = &chain.extension_list_xml {
        if extension.len() > MAX_EXTENSION_BYTES {
            return Err(invalid("calculation-chain extension list is too large"));
        }
        add_len(&mut len, extension.len())?;
    }
    add_len(&mut len, "</calcChain>".len())?;
    if len > MAX_OUTPUT_BYTES {
        return Err(invalid(format!(
            "calculation-chain output exceeds {MAX_OUTPUT_BYTES} bytes"
        )));
    }
    Ok(len)
}

fn attribute_len(name: &str, value: &str) -> Result<usize> {
    validate_attribute_size(name, value)?;
    let mut len = name.len();
    add_len(&mut len, escaped_len(value)?)?;
    add_len(&mut len, 4)?;
    Ok(len)
}

fn escaped_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |mut len, character| {
        let bytes = match character {
            '&' => 5,
            '<' => 4,
            '"' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        add_len(&mut len, bytes)?;
        Ok(len)
    })
}

fn add_len(total: &mut usize, value: usize) -> Result<()> {
    *total = total
        .checked_add(value)
        .ok_or_else(|| invalid("calculation-chain output length overflow"))?;
    Ok(())
}

fn decimal_len(mut value: u32) -> usize {
    let mut len = 1;
    while value >= 10 {
        value /= 10;
        len += 1;
    }
    len
}

fn push_u16(output: &mut String, mut value: u16) {
    let mut digits = [0u8; 5];
    let mut start = digits.len();
    loop {
        start -= 1;
        digits[start] = b'0' + (value % 10) as u8;
        value /= 10;
        if value == 0 {
            break;
        }
    }
    for digit in &digits[start..] {
        output.push(char::from(*digit));
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

fn enforce_budget(start: usize, current: usize, limit: usize, message: &'static str) -> Result<()> {
    let consumed = current
        .checked_sub(start)
        .ok_or_else(|| invalid("calculation-chain XML offset moved backwards"))?;
    if consumed > limit {
        return Err(invalid(message));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::part::{BlobPart, Part};

    #[test]
    fn parses_writes_typed_sheets_steps_strict_and_extensions() {
        let xml = br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="urn:test" x:root="v"><c r="A1" i="2" l="1" t="true" a="1" x:cell="kept"/><c r="B2" s="1"/><c r="C3"/><extLst><ext uri="urn:test"><x:data value="inert"/></ext></extLst></calcChain>"#;
        let chain = read(xml).unwrap();
        assert_eq!(chain.len(), 3);
        assert_eq!(chain.at(1).unwrap().sheet().get(), 2);
        assert_eq!(chain.at(0).unwrap().step(), Step::Level);
        assert_eq!(chain.at(1).unwrap().step(), Step::Child);
        assert_eq!(chain.at(2).unwrap().step(), Step::Same);
        assert!(chain.at(0).unwrap().flags().contains(Flags::THREAD));
        assert!(chain.at(0).unwrap().flags().contains(Flags::ARRAY));
        assert_eq!(
            chain.get(Sheet::new(2).unwrap(), "B2").unwrap(),
            chain.cells().get(1)
        );

        let strict = String::from_utf8(write(&chain, Conformance::Strict).unwrap()).unwrap();
        assert!(strict.contains(STRICT_NS));
        assert!(strict.contains("x:cell=\"kept\""));
        assert!(strict.contains("<extLst>"));
        let reparsed = read(strict.as_bytes()).unwrap();
        assert_eq!(reparsed, chain);
        assert_eq!(
            write(&reparsed, Conformance::Strict).unwrap(),
            strict.as_bytes()
        );
    }

    #[test]
    fn preprocesses_mce_and_rejects_malformed_records() {
        let mce = br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:c/></mc:Choice><mc:Fallback><c r="C3" i="1"/></mc:Fallback></mc:AlternateContent></calcChain>"#;
        assert_eq!(read(mce).unwrap().cells()[0].reference(), "C3");
        let invalid = [
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"/>"#),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c/></calcChain>"#),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1"/></calcChain>"#),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="XFE1" i="1"/></calcChain>"#),
            format!(
                r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1" l="yes"/></calcChain>"#
            ),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="65535"/></calcChain>"#),
            format!(
                r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1" l="1" s="1"/></calcChain>"#
            ),
            format!(
                r#"<calcChain xmlns="{TRANSITIONAL_NS}"><extLst/><c r="A1" i="1"/></calcChain>"#
            ),
            format!(
                r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"><c r="B1"/></c></calcChain>"#
            ),
            format!(
                r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1" bogus="1"/></calcChain>"#
            ),
        ];
        for xml in invalid {
            assert!(read(xml.as_bytes()).is_err(), "accepted {xml}");
        }
        assert!(Sheet::new(0).is_err());
        assert!(Sheet::new(65_535).is_err());
        assert_eq!(Sheet::new(65_534).unwrap().get(), 65_534);
    }

    #[test]
    fn semantic_and_positional_crud_is_checked_and_failure_atomic() {
        let sheet = Sheet::new(1).unwrap();
        let first = Cell::new(sheet, "A1").unwrap();
        let mut chain = Chain::new(first.clone());
        chain.push(Cell::new(sheet, "C3").unwrap()).unwrap();
        chain.insert(1, Cell::new(sheet, "B2").unwrap()).unwrap();
        assert_eq!(chain.at(1).unwrap().reference(), "B2");
        assert_eq!(
            chain.get(sheet, "C3").unwrap().unwrap().address(),
            Address::from_a1("C3").unwrap()
        );

        let before = chain.clone();
        assert!(chain.push(Cell::new(sheet, "B2").unwrap()).is_err());
        assert!(chain.insert(9, Cell::new(sheet, "D4").unwrap()).is_err());
        assert_eq!(chain, before);

        let replaced = chain
            .replace_at(1, Cell::new(sheet, "D4").unwrap())
            .unwrap();
        assert_eq!(replaced.reference(), "B2");
        let replaced = chain.put(Cell::new(sheet, "D4").unwrap()).unwrap().unwrap();
        assert_eq!(replaced.reference(), "D4");
        assert_eq!(
            chain.remove(sheet, "C3").unwrap().unwrap().reference(),
            "C3"
        );
        chain.move_at(1, 0).unwrap();
        assert_eq!(chain.at(0).unwrap().reference(), "D4");
        assert_eq!(chain.remove_at(1).unwrap().reference(), "A1");
        let before = chain.clone();
        assert!(chain.remove_at(0).is_err());
        assert_eq!(chain, before);
        assert!(!chain.is_empty());
    }

    #[test]
    fn malformed_duplicate_keys_are_retained_but_never_selected_silently() {
        let xml = format!(
            r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"/><c r="A1"/></calcChain>"#
        );
        let chain = read(xml.as_bytes()).unwrap();
        assert_eq!(chain.len(), 2);
        assert!(chain.get(Sheet::new(1).unwrap(), "A1").is_err());
    }

    #[test]
    fn semantic_mutations_reject_ambiguous_imports_without_changing_order() {
        let xml = format!(
            r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"/><c r="A1"/><c r="B2"/></calcChain>"#
        );
        let mut chain = read(xml.as_bytes()).unwrap();
        let before = chain.clone();
        let sheet = Sheet::new(1).unwrap();

        assert!(chain.get(sheet, "B2").is_err());
        assert!(chain.put(Cell::new(sheet, "C3").unwrap()).is_err());
        assert!(chain.push(Cell::new(sheet, "C3").unwrap()).is_err());
        assert!(chain.insert(1, Cell::new(sheet, "C3").unwrap()).is_err());
        assert!(chain.remove(sheet, "B2").is_err());
        assert_eq!(chain, before);
    }

    #[test]
    fn positional_repairs_refresh_ambiguous_import_state() {
        let xml = format!(
            r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1"/><c r="A1"/><c r="B2"/></calcChain>"#
        );
        let mut chain = read(xml.as_bytes()).unwrap();
        let sheet = Sheet::new(1).unwrap();

        chain.remove_at(1).unwrap();
        assert_eq!(chain.get(sheet, "B2").unwrap().unwrap().reference(), "B2");

        let mut chain = read(xml.as_bytes()).unwrap();
        chain
            .replace_at(1, Cell::new(sheet, "C3").unwrap())
            .unwrap();
        assert_eq!(chain.get(sheet, "C3").unwrap().unwrap().reference(), "C3");
    }

    #[test]
    fn rejects_oversized_nested_calculation_content_before_decoding() {
        let mut xml =
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="A1" i="1">"#).into_bytes();
        xml.extend(std::iter::repeat_n(
            b' ',
            MAX_CELL_CONTENT_BYTES.saturating_add(1),
        ));
        xml.extend_from_slice(b"</c></calcChain>");

        assert!(read(&xml).is_err());
    }

    #[test]
    fn adversarial_xml_returns_errors_without_unwinding() {
        let inputs: [&[u8]; 4] = [
            b"<",
            b"\xff<calcChain/>",
            br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="0"/></calcChain>"#,
            br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1" \xff="bad"/></calcChain>"#,
        ];
        for input in inputs {
            let result = std::panic::catch_unwind(|| read(input));
            assert!(result.is_ok(), "reader unwound for {input:?}");
            assert!(result.unwrap().is_err());
        }
    }

    #[test]
    fn stores_rewrites_and_removes_inert_calculation_chain_parts() {
        let mut package = workbook_package();
        let mut first = Cell::new(Sheet::new(1).unwrap(), "B2").unwrap();
        first.set_step(Step::Level);
        let chain = Chain::new(first);

        assert!(put(&mut package, &chain, Conformance::Transitional).unwrap());
        assert_eq!(
            load(&package).unwrap(),
            Some((chain.clone(), Conformance::Transitional))
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
        let before = package.get_part(&part_name).unwrap().blob_arc();
        assert!(!put(&mut package, &chain, Conformance::Transitional).unwrap());
        let after = package.get_part(&part_name).unwrap().blob_arc();
        assert!(std::sync::Arc::ptr_eq(&before, &after));

        let replacement = Chain::new(Cell::new(Sheet::new(1).unwrap(), "C3").unwrap());
        assert!(put(&mut package, &replacement, Conformance::Strict).unwrap());
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
            load(&package).unwrap(),
            Some((replacement, Conformance::Strict))
        );

        assert!(remove(&mut package).unwrap());
        assert!(package.get_part(&part_name).is_err());
        assert_eq!(load(&package).unwrap(), None);
        assert!(!remove(&mut package).unwrap());
    }

    #[test]
    fn removal_retains_a_calculation_chain_part_referenced_elsewhere() {
        let mut package = workbook_package();
        let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "F6").unwrap());
        put(&mut package, &chain, Conformance::Transitional).unwrap();

        let part_name = PackURI::new("/xl/calcChain.xml").unwrap();
        let mut referring_part = BlobPart::new(
            PackURI::new("/xl/retained-reference.xml").unwrap(),
            ct::XML.into(),
            b"<reference/>".to_vec(),
        );
        referring_part.relate_to("calcChain.xml", "urn:litchi:test:calc-chain-reference");
        package.add_part(Box::new(referring_part));

        assert!(remove(&mut package).unwrap());
        assert!(package.get_part(&part_name).is_ok());
        assert!(load(&package).is_err());
        assert!(put(&mut package, &chain, Conformance::Transitional).is_err());
    }

    #[test]
    fn package_calculation_chain_mutators_reject_invalid_existing_graphs() {
        let mut package = synthetic_package(RELATIONSHIP, false, ct::XML, false);
        let chain_part = PackURI::new("/xl/calcChain.xml").unwrap();
        let original = package.get_part(&chain_part).unwrap().blob().to_vec();
        let chain = Chain::new(Cell::new(Sheet::new(1).unwrap(), "E5").unwrap());

        assert!(put(&mut package, &chain, Conformance::Transitional).is_err());
        assert_eq!(package.get_part(&chain_part).unwrap().blob(), original);
        assert!(remove(&mut package).is_err());
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
        assert!(put(&mut duplicate, &chain, Conformance::Transitional).is_err());
        assert!(remove(&mut duplicate).is_err());

        let mut duplicate_part = synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, false);
        duplicate_part.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/calcChainExtra.xml").unwrap(),
            CONTENT_TYPE.into(),
            format!(r#"<calcChain xmlns="{TRANSITIONAL_NS}"><c r="F6"/></calcChain>"#).into_bytes(),
        )));
        assert!(load(&duplicate_part).is_err());
        assert!(put(&mut duplicate_part, &chain, Conformance::Transitional).is_err());
        assert!(remove(&mut duplicate_part).is_err());

        let mut external = synthetic_package(RELATIONSHIP, true, CONTENT_TYPE, false);
        assert!(put(&mut external, &chain, Conformance::Transitional).is_err());
        assert!(remove(&mut external).is_err());
    }

    #[test]
    fn loads_real_poi_and_synthetic_packages_and_validates_relationships() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..//test-data/poi/test-data/spreadsheet/62834.xlsx");
        let package = OpcPackage::open(path).unwrap();
        let (chain, _) = load(&package).unwrap().unwrap();
        assert_eq!(chain.cells().len(), 3);
        assert_eq!(chain.cells()[0].reference(), "A5");
        assert_eq!(chain.cells()[0].step(), Step::Level);
        assert_eq!(chain.cells()[2].step(), Step::Child);

        let package = synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, false);
        assert_eq!(
            load(&package).unwrap().unwrap().0.cells()[0].reference(),
            "A1"
        );

        assert!(load(&synthetic_package(RELATIONSHIP, true, CONTENT_TYPE, false)).is_err());
        assert!(load(&synthetic_package(RELATIONSHIP, false, ct::XML, false)).is_err());
        assert!(load(&synthetic_package(RELATIONSHIP, false, CONTENT_TYPE, true)).is_err());
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
