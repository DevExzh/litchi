//! Typed calculation-chain semantic models.

use std::collections::HashSet;

use crate::error::{Error, Result, allocation, invalid};
use litchi_sheet::{At, Cell as Address};

pub(crate) const TRANSITIONAL_NS: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT_NS: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
pub(crate) const RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
pub(crate) const STRICT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain";
pub(crate) const MAX_CELLS: usize = 2_000_000;
pub(crate) const MAX_XML_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_EXTENSION_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_CELL_CONTENT_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_EXTENSION_ATTRIBUTES: usize = 256;
pub(crate) const MAX_ATTRIBUTE_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_EXTENSION_DEPTH: usize = 128;
pub(crate) const MAX_REFERENCE_BYTES: usize = 32;

/// Namespace family used by the calculation-chain writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Conformance {
    #[default]
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) const fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_NS,
            Self::Strict => STRICT_NS,
        }
    }

    pub(crate) const fn relationship_type(self) -> &'static str {
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
        pub(crate) name: String,
        pub(crate) value: String,
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
    pub(crate) reference: Box<str>,
    pub(crate) address: Address,
    pub(crate) sheet: Sheet,
    pub(crate) explicit_sheet: bool,
    pub(crate) step: Step,
    pub(crate) flags: Flags,
    pub(crate) attrs: Vec<Attr>,
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
    pub(crate) cells: Vec<Cell>,
    pub(crate) ambiguous_key: Option<(Sheet, Address)>,
    pub(crate) extension_list_xml: Option<String>,
    pub(crate) namespace_declarations: Vec<(String, String)>,
    pub(crate) attrs: Vec<Attr>,
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

    pub(crate) fn ensure_sheet_boundaries(&mut self) {
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
