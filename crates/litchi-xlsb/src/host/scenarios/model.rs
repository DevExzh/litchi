//! Typed XLSB Scenario Manager values.
//!
//! The model follows `[MS-XLSB]` 2.4.198, 2.4.199, and 2.4.797.  It is an
//! inert snapshot: changed-cell values are metadata for what-if analysis and
//! are never applied to worksheet cells or recalculated by litchi.

use super::validation;
use crate::package::error::Result;

/// Maximum number of changing cells in one scenario.
pub const MAX_CHANGED_CELLS: usize = 32;
/// Maximum number of result ranges in one Scenario Manager.
pub const MAX_RESULT_RANGES: usize = 32;
/// Maximum number of scenarios retained by one manager.
pub const MAX_SCENARIOS: usize = 65_535;
/// Maximum UTF-16 code units in a scenario name or comment.
pub const MAX_SCENARIO_TEXT: usize = 255;
/// Maximum UTF-16 code units in the user name field.
pub const MAX_USER_NAME: usize = 54;
/// Maximum bytes retained in one opaque record payload.
pub const MAX_UNKNOWN_PAYLOAD: usize = 64 * 1024 * 1024;
/// Maximum opaque record count in one scenario collection.
pub const MAX_UNKNOWN_RECORDS: usize = 65_536;

/// An inclusive zero-based worksheet cell range (`UncheckedRfX`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRange {
    row_first: u32,
    row_last: u32,
    column_first: u32,
    column_last: u32,
}

impl CellRange {
    /// Construct a checked inclusive worksheet range.
    pub fn new(row_first: u32, row_last: u32, column_first: u32, column_last: u32) -> Result<Self> {
        let value = Self {
            row_first,
            row_last,
            column_first,
            column_last,
        };
        validation::validate_range(value)?;
        Ok(value)
    }

    /// Construct a one-cell range.
    pub fn cell(row: u32, column: u32) -> Result<Self> {
        Self::new(row, row, column, column)
    }

    #[must_use]
    pub const fn row_first(&self) -> u32 {
        self.row_first
    }

    #[must_use]
    pub const fn row_last(&self) -> u32 {
        self.row_last
    }

    #[must_use]
    pub const fn column_first(&self) -> u32 {
        self.column_first
    }

    #[must_use]
    pub const fn column_last(&self) -> u32 {
        self.column_last
    }

    /// Number of cells covered by this range.
    #[must_use]
    pub fn cell_count(&self) -> u64 {
        u64::from(self.row_last - self.row_first + 1)
            * u64::from(self.column_last - self.column_first + 1)
    }
}

/// A changing cell (`BrtSlc`) belonging to a scenario.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChangedCell {
    row: u32,
    column: u32,
    number_format: u16,
    value: String,
    /// `fUnused` is ignored by Excel but retained for lossless snapshots.
    unused: u32,
}

impl ChangedCell {
    /// Construct a changed cell with the default number format.
    pub fn new(row: u32, column: u32, value: impl Into<String>) -> Result<Self> {
        Self::with_number_format(row, column, 0, value)
    }

    /// Construct a changed cell with its scenario display format.
    pub fn with_number_format(
        row: u32,
        column: u32,
        number_format: u16,
        value: impl Into<String>,
    ) -> Result<Self> {
        let value = Self {
            row,
            column,
            number_format,
            value: value.into(),
            unused: 0,
        };
        validation::validate_changed_cell(&value)?;
        Ok(value)
    }

    #[must_use]
    pub const fn row(&self) -> u32 {
        self.row
    }

    #[must_use]
    pub const fn column(&self) -> u32 {
        self.column
    }

    #[must_use]
    pub const fn number_format(&self) -> u16 {
        self.number_format
    }

    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Replace the displayed value after applying the XLSB text bound.
    pub fn set_value(&mut self, value: impl Into<String>) -> Result<()> {
        let previous = std::mem::replace(&mut self.value, value.into());
        if let Err(error) = validation::validate_changed_cell(self) {
            self.value = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Replace the number format identifier.
    pub const fn set_number_format(&mut self, number_format: u16) {
        self.number_format = number_format;
    }

    /// The ignored wire `fUnused` value retained for lossless edits.
    #[must_use]
    pub const fn unused(&self) -> u32 {
        self.unused
    }

    pub(crate) fn from_wire(
        row: u32,
        column: u32,
        number_format: u16,
        value: String,
        unused: u32,
    ) -> Result<Self> {
        let cell = Self {
            row,
            column,
            number_format,
            value,
            unused,
        };
        validation::validate_changed_cell(&cell)?;
        Ok(cell)
    }
}

/// One unrecognised BIFF12 record retained as inert data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownRecord {
    kind: u16,
    payload: Box<[u8]>,
}

impl UnknownRecord {
    #[must_use]
    pub const fn kind(&self) -> u16 {
        self.kind
    }

    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    pub(crate) fn new(kind: u16, payload: &[u8]) -> Result<Self> {
        if kind > 0x3fff {
            return Err(validation::invalid(
                "unknown scenario record kind",
                format!("0x{kind:04X}"),
            ));
        }
        if payload.len() > MAX_UNKNOWN_PAYLOAD {
            return Err(validation::invalid(
                "unknown scenario record",
                format!("payload exceeds {MAX_UNKNOWN_PAYLOAD} bytes"),
            ));
        }
        Ok(Self {
            kind,
            payload: payload.to_vec().into_boxed_slice(),
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Child {
    Changed(usize),
    Unknown(usize),
}

/// One named what-if scenario (`BrtBeginSct` plus its `BrtSlc` records).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Scenario {
    name: String,
    locked: bool,
    hidden: bool,
    comment: String,
    user_name: String,
    changed_cells: Vec<ChangedCell>,
    unknown: Vec<UnknownRecord>,
    pub(crate) order: Vec<Child>,
}

impl Scenario {
    /// Create a scenario.  A scenario must be given at least one changed
    /// cell before it can be serialized.
    pub fn new(name: impl Into<String>, user_name: impl Into<String>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            locked: false,
            hidden: false,
            comment: String::new(),
            user_name: user_name.into(),
            changed_cells: Vec::new(),
            unknown: Vec::new(),
            order: Vec::new(),
        };
        validation::validate_scenario_text(&value)?;
        Ok(value)
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub const fn locked(&self) -> bool {
        self.locked
    }

    #[must_use]
    pub const fn hidden(&self) -> bool {
        self.hidden
    }

    #[must_use]
    pub fn comment(&self) -> &str {
        &self.comment
    }

    #[must_use]
    pub fn user_name(&self) -> &str {
        &self.user_name
    }

    #[must_use]
    pub fn changed_cells(&self) -> &[ChangedCell] {
        &self.changed_cells
    }

    #[must_use]
    pub fn unknown_records(&self) -> &[UnknownRecord] {
        &self.unknown
    }

    pub fn set_name(&mut self, name: impl Into<String>) -> Result<()> {
        let previous = std::mem::replace(&mut self.name, name.into());
        if let Err(error) = validation::validate_scenario_text(self) {
            self.name = previous;
            return Err(error);
        }
        Ok(())
    }

    pub fn set_comment(&mut self, comment: impl Into<String>) -> Result<()> {
        let previous = std::mem::replace(&mut self.comment, comment.into());
        if let Err(error) = validation::validate_scenario_text(self) {
            self.comment = previous;
            return Err(error);
        }
        Ok(())
    }

    pub fn set_user_name(&mut self, user_name: impl Into<String>) -> Result<()> {
        let previous = std::mem::replace(&mut self.user_name, user_name.into());
        if let Err(error) = validation::validate_scenario_text(self) {
            self.user_name = previous;
            return Err(error);
        }
        Ok(())
    }

    pub const fn set_locked(&mut self, locked: bool) {
        self.locked = locked;
    }

    pub const fn set_hidden(&mut self, hidden: bool) {
        self.hidden = hidden;
    }

    /// Replace changing cells while retaining opaque records when the wire
    /// cardinality remains unchanged. Structural edits beside unknown records
    /// are refused because their safe insertion point cannot be inferred.
    pub fn set_changed_cells(&mut self, changed_cells: Vec<ChangedCell>) -> Result<()> {
        if !self.unknown.is_empty() && changed_cells.len() != self.changed_cells.len() {
            return Err(validation::invalid(
                "scenario edit",
                "cannot change BrtSlc cardinality while unknown records are present",
            ));
        }
        let previous_order = std::mem::take(&mut self.order);
        let previous = std::mem::replace(&mut self.changed_cells, changed_cells);
        if !self.unknown.is_empty() {
            self.order = previous_order.clone();
        }
        if let Err(error) = self.validate() {
            self.changed_cells = previous;
            self.order = previous_order;
            return Err(error);
        }
        if self.unknown.is_empty() {
            self.order = (0..self.changed_cells.len()).map(Child::Changed).collect();
        }
        Ok(())
    }

    /// Replace one changing cell without changing the record order.
    pub fn set_changed_cell(&mut self, index: usize, cell: ChangedCell) -> Result<()> {
        if self.changed_cells.get(index).is_none() {
            return Err(validation::invalid(
                "scenario edit",
                format!("changed-cell index {index} is out of bounds"),
            ));
        }
        let previous = std::mem::replace(&mut self.changed_cells[index], cell);
        if let Err(error) = self.validate() {
            self.changed_cells[index] = previous;
            return Err(error);
        }
        Ok(())
    }

    pub(crate) fn from_wire(
        name: String,
        locked: bool,
        hidden: bool,
        comment: String,
        user_name: String,
        changed_cells: Vec<ChangedCell>,
        unknown: Vec<UnknownRecord>,
        order: Vec<Child>,
    ) -> Result<Self> {
        let value = Self {
            name,
            locked,
            hidden,
            comment,
            user_name,
            changed_cells,
            unknown,
            order,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validation::validate_scenario(self)
    }

    pub(crate) fn is_opaque(&self) -> bool {
        !self.unknown.is_empty()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Entry {
    Scenario(usize),
    Unknown(usize),
}

/// The worksheet Scenario Manager (`BrtBeginScenMan` collection).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Manager {
    current: Option<usize>,
    shown: Option<usize>,
    result_ranges: Vec<CellRange>,
    scenarios: Vec<Scenario>,
    unknown: Vec<UnknownRecord>,
    pub(crate) order: Vec<Entry>,
}

impl Manager {
    /// Construct a Scenario Manager from scenarios in source order.
    pub fn new(scenarios: Vec<Scenario>) -> Result<Self> {
        let order = (0..scenarios.len()).map(Entry::Scenario).collect();
        let value = Self {
            current: None,
            shown: None,
            result_ranges: Vec::new(),
            scenarios,
            unknown: Vec::new(),
            order,
        };
        value.validate()?;
        Ok(value)
    }

    #[must_use]
    pub const fn current(&self) -> Option<usize> {
        self.current
    }

    #[must_use]
    pub const fn shown(&self) -> Option<usize> {
        self.shown
    }

    #[must_use]
    pub fn result_ranges(&self) -> &[CellRange] {
        &self.result_ranges
    }

    #[must_use]
    pub fn scenarios(&self) -> &[Scenario] {
        &self.scenarios
    }

    /// The currently selected scenario, if one is selected.
    #[must_use]
    pub fn current_scenario(&self) -> Option<&Scenario> {
        self.current.and_then(|index| self.scenarios.get(index))
    }

    /// The last shown scenario, if one is recorded.
    #[must_use]
    pub fn shown_scenario(&self) -> Option<&Scenario> {
        self.shown.and_then(|index| self.scenarios.get(index))
    }

    #[must_use]
    pub fn unknown_records(&self) -> &[UnknownRecord] {
        &self.unknown
    }

    pub fn set_current(&mut self, current: Option<usize>) -> Result<()> {
        let previous = self.current;
        self.current = current;
        if let Err(error) = self.validate() {
            self.current = previous;
            return Err(error);
        }
        Ok(())
    }

    pub fn set_shown(&mut self, shown: Option<usize>) -> Result<()> {
        let previous = self.shown;
        self.shown = shown;
        if let Err(error) = self.validate() {
            self.shown = previous;
            return Err(error);
        }
        Ok(())
    }

    pub fn set_result_ranges(&mut self, result_ranges: Vec<CellRange>) -> Result<()> {
        let previous = std::mem::replace(&mut self.result_ranges, result_ranges);
        if let Err(error) = self.validate() {
            self.result_ranges = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Replace one scenario without changing the manager record order.
    pub fn set_scenario(&mut self, index: usize, scenario: Scenario) -> Result<()> {
        if self.scenarios.get(index).is_none() {
            return Err(validation::invalid(
                "scenario edit",
                format!("scenario index {index} is out of bounds"),
            ));
        }
        let previous = std::mem::replace(&mut self.scenarios[index], scenario);
        if let Err(error) = self.validate() {
            self.scenarios[index] = previous;
            return Err(error);
        }
        Ok(())
    }

    /// Replace all scenarios. Structural changes beside opaque records are
    /// rejected because no safe placement for new records can be inferred.
    pub fn set_scenarios(&mut self, scenarios: Vec<Scenario>) -> Result<()> {
        if !self.unknown.is_empty() && scenarios.len() != self.scenarios.len() {
            return Err(validation::invalid(
                "scenario edit",
                "cannot change scenario cardinality while unknown records are present",
            ));
        }
        let previous_order = std::mem::take(&mut self.order);
        let previous = std::mem::replace(&mut self.scenarios, scenarios);
        if !self.unknown.is_empty() {
            self.order = previous_order.clone();
        }
        if let Err(error) = self.validate() {
            self.scenarios = previous;
            self.order = previous_order;
            return Err(error);
        }
        if self.unknown.is_empty() {
            self.order = (0..self.scenarios.len()).map(Entry::Scenario).collect();
        }
        Ok(())
    }

    pub(crate) fn from_wire(
        current: Option<usize>,
        shown: Option<usize>,
        result_ranges: Vec<CellRange>,
        scenarios: Vec<Scenario>,
        unknown: Vec<UnknownRecord>,
        order: Vec<Entry>,
    ) -> Result<Self> {
        let value = Self {
            current,
            shown,
            result_ranges,
            scenarios,
            unknown,
            order,
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validation::validate_manager(self)
    }

    pub(crate) fn is_opaque(&self) -> bool {
        !self.unknown.is_empty() || self.scenarios.iter().any(Scenario::is_opaque)
    }
}
