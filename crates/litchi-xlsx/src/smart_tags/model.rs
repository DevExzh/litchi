use litchi_sheet::{At, Cell as Address};

use crate::error::Result;

/// `SpreadsheetML` namespace form used by canonical serialization.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum Conformance {
    #[default]
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) const fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => super::TRANSITIONAL,
            Self::Strict => super::STRICT,
        }
    }
}

/// One inert key/value property attached to a tag.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Property {
    pub(crate) key: Box<str>,
    pub(crate) value: Box<str>,
}

impl Property {
    /// Create a validated property.
    pub fn new(key: impl Into<Box<str>>, value: impl Into<Box<str>>) -> Result<Self> {
        let value = Self {
            key: key.into(),
            value: value.into(),
        };
        super::validation::property(&value)?;
        Ok(value)
    }

    #[must_use]
    pub fn key(&self) -> &str {
        &self.key
    }

    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// One typed smart-tag annotation on a cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Tag {
    pub(crate) type_id: u32,
    pub(crate) deleted: bool,
    pub(crate) xml_based: bool,
    pub(crate) properties: Vec<Property>,
}

impl Tag {
    /// Create a tag using an Office smart-tag type identifier.
    pub fn new(type_id: u32) -> Result<Self> {
        super::validation::type_id(type_id)?;
        Ok(Self {
            type_id,
            deleted: false,
            xml_based: false,
            properties: Vec::new(),
        })
    }

    #[must_use]
    pub const fn type_id(&self) -> u32 {
        self.type_id
    }

    #[must_use]
    pub const fn is_deleted(&self) -> bool {
        self.deleted
    }

    #[must_use]
    pub const fn is_xml_based(&self) -> bool {
        self.xml_based
    }

    #[must_use]
    pub fn properties(&self) -> &[Property] {
        &self.properties
    }

    pub fn set_deleted(&mut self, value: bool) -> &mut Self {
        self.deleted = value;
        self
    }

    pub fn set_xml_based(&mut self, value: bool) -> &mut Self {
        self.xml_based = value;
        self
    }

    /// Add a property after checking collection-wide key uniqueness.
    pub fn add_property(&mut self, property: Property) -> Result<&mut Self> {
        if self
            .properties
            .iter()
            .any(|candidate| candidate.key == property.key)
        {
            return Err(crate::error::invalid(format!(
                "duplicate smart-tag property key '{}'",
                property.key
            )));
        }
        self.properties.push(property);
        if let Err(error) = super::validation::tag(self) {
            self.properties.pop();
            return Err(error);
        }
        Ok(self)
    }
}

/// All smart tags attached to one checked worksheet cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cell {
    pub(crate) address: Address,
    pub(crate) tags: Vec<Tag>,
}

impl Cell {
    /// Create a non-empty cell annotation.
    pub fn new<'a>(at: impl Into<At<'a>>, tags: Vec<Tag>) -> Result<Self> {
        let value = Self {
            address: at.into().resolve()?,
            tags,
        };
        super::validation::cell(&value)?;
        Ok(value)
    }

    #[must_use]
    pub const fn address(&self) -> Address {
        self.address
    }

    #[must_use]
    pub fn tags(&self) -> &[Tag] {
        &self.tags
    }

    /// Add another tag to this cell.
    pub fn add(&mut self, tag: Tag) -> Result<&mut Self> {
        self.tags.push(tag);
        if let Err(error) = super::validation::cell(self) {
            self.tags.pop();
            return Err(error);
        }
        Ok(self)
    }
}

/// Worksheet smart tags in deterministic cell order.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Collection {
    cells: Vec<Cell>,
}

impl Collection {
    pub fn new(mut cells: Vec<Cell>) -> Result<Self> {
        cells.sort_unstable_by_key(Cell::address);
        let value = Self { cells };
        super::validation::collection(&value)?;
        Ok(value)
    }

    #[must_use]
    pub fn cells(&self) -> &[Cell] {
        &self.cells
    }

    pub fn get<'a>(&self, at: impl Into<At<'a>>) -> Result<Option<&Cell>> {
        let address = at.into().resolve()?;
        Ok(self
            .cells
            .binary_search_by_key(&address, Cell::address)
            .ok()
            .map(|index| &self.cells[index]))
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.cells.len()
    }

    pub(crate) fn upsert(&mut self, value: Cell) {
        match self
            .cells
            .binary_search_by_key(&value.address(), Cell::address)
        {
            Ok(index) => self.cells[index] = value,
            Err(index) => self.cells.insert(index, value),
        }
    }

    pub(crate) fn remove(&mut self, address: Address) -> Option<Cell> {
        self.cells
            .binary_search_by_key(&address, Cell::address)
            .ok()
            .map(|index| self.cells.remove(index))
    }
}
