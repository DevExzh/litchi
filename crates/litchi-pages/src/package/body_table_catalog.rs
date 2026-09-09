//! Bounded, selector-first Pages body-table discovery.
//!
//! The catalog is a read-only semantic projection.  Native attachments,
//! archive members, generated messages, and object identifiers remain inside
//! the package owner; callers receive only source position, validated name,
//! and checked dimensions.

use std::fmt;

use thiserror::Error;

use crate::Position;
use crate::package::{Package, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::name::{Error as NameError, Name};

/// Finite resources charged while discovering one immutable body-table
/// catalog.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableCatalogLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete package output bytes (reserved for shared package budgets).
    OutputBytes,
    /// ZIP members retained by the package.
    Entries,
    /// Bytes retained by one ZIP member.
    EntryBytes,
    /// Aggregate bytes retained by ZIP members.
    TotalEntryBytes,
    /// ZIP container names or structural metadata bytes.
    PackageBytes,
    /// Bytes in one decoded native payload container.
    PayloadBytes,
    /// Aggregate decoded native payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected by discovery.
    PayloadObjects,
    /// Native payload messages inspected by discovery.
    PayloadMessages,
    /// Native payload framing or metadata items inspected by discovery.
    PayloadItems,
    /// Native object references inspected while proving ownership.
    PayloadReferences,
    /// Strict protobuf bytes.
    WireBytes,
    /// Strict protobuf fields.
    WireFields,
    /// Strict protobuf nesting.
    WireNesting,
    /// Aggregate wire and graph work.
    WireWork,
}

impl fmt::Display for BodyTableCatalogLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "ZIP entries",
            Self::EntryBytes => "ZIP entry bytes",
            Self::TotalEntryBytes => "total ZIP entry bytes",
            Self::PackageBytes => "package metadata bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// Failure while discovering or selecting a Pages body-table catalog.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableCatalogError {
    /// More than one rooted body table has the requested exact name.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The rooted native graph or selected table payload is malformed.
    #[error("the Pages body-table catalog source is invalid")]
    InvalidSource,
    /// A discovered table name violates the public semantic name invariant.
    #[error("invalid Pages body-table name: {0}")]
    InvalidName(NameError),
    /// A finite catalog resource ceiling was exceeded.
    #[error(
        "Pages body-table catalog {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: BodyTableCatalogLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded catalog allocation failed.
    #[error("could not allocate {amount} units for the Pages body-table catalog")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
    },
}

/// One immutable semantic body-table snapshot.
///
/// The snapshot owns only validated display text and dimensions.  Its
/// position is the zero-based order among rooted body tables in the package;
/// it is not a native object identifier.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct BodyTableSnapshot {
    position: Position,
    name: Name,
    rows: u32,
    columns: u32,
}

impl BodyTableSnapshot {
    fn from_target(entry: table_lock::BodyTableTarget) -> Result<Self, BodyTableCatalogError> {
        let name = Name::try_from(entry.table_name.into_string())
            .map_err(BodyTableCatalogError::InvalidName)?;
        Ok(Self {
            position: Position::new(entry.table_position),
            name,
            rows: entry.table_rows,
            columns: entry.table_columns,
        })
    }

    /// Return the checked zero-based source position of this table.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the zero-based source index of this table.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.position.get()
    }

    /// Borrow the validated table name.
    #[must_use]
    pub fn name(&self) -> &str {
        self.name.as_str()
    }

    /// Borrow the validated owned table-name value.
    #[must_use]
    pub const fn name_value(&self) -> &Name {
        &self.name
    }

    /// Return the declared row count.
    #[must_use]
    pub const fn rows(&self) -> u32 {
        self.rows
    }

    /// Return the declared column count.
    #[must_use]
    pub const fn columns(&self) -> u32 {
        self.columns
    }

    /// Return a position selector for this snapshot.
    ///
    /// The position is the stable selector within this immutable catalog
    /// snapshot and remains usable when malformed input repeats a name. It is
    /// a snapshot-local coordinate, not a durable identity across packages.
    #[must_use]
    pub const fn selector(&self) -> BodyTableSelector<'static> {
        BodyTableSelector::position(self.position)
    }

    /// Return an exact-name selector for this snapshot.
    #[must_use]
    pub fn name_selector(&self) -> BodyTableSelector<'_> {
        BodyTableSelector::name(self.name())
    }
}

/// Immutable, bounded semantic body-table catalog in rooted source order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BodyTableCatalog {
    values: Vec<BodyTableSnapshot>,
}

impl BodyTableCatalog {
    fn from_targets(
        entries: Vec<table_lock::BodyTableTarget>,
        budget: &mut table_lock::WireBudget,
    ) -> Result<Self, BodyTableCatalogError> {
        budget
            .charge_payload_items(entries.len())
            .map_err(map_lock_error)?;
        for entry in &entries {
            budget
                .charge_payload_work(entry.table_name.len())
                .map_err(map_lock_error)?;
        }
        let mut values = Vec::new();
        values
            .try_reserve_exact(entries.len())
            .map_err(|_| BodyTableCatalogError::Allocation {
                amount: entries.len(),
            })?;
        for entry in entries {
            values.push(BodyTableSnapshot::from_target(entry)?);
        }
        Ok(Self { values })
    }

    /// Return the number of rooted body tables.
    #[must_use]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    /// Return whether the rooted body has no tables.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    /// Borrow all semantic table snapshots in rooted source order.
    #[must_use]
    pub fn as_slice(&self) -> &[BodyTableSnapshot] {
        &self.values
    }

    /// Borrow one table by zero-based rooted source order.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&BodyTableSnapshot> {
        self.values.get(index)
    }

    /// Resolve a name- or position-based selector.
    ///
    /// Position selectors return `None` for an out-of-range index.  Name
    /// selectors return `None` for a missing name and an error when malformed
    /// input exposes duplicate exact names, matching the existing Pages
    /// selector policy.
    pub fn select<'selector, S>(
        &self,
        selector: S,
    ) -> Result<Option<&BodyTableSnapshot>, BodyTableCatalogError>
    where
        S: Into<BodyTableSelector<'selector>>,
    {
        match selector.into() {
            BodyTableSelector::Position(position) => Ok(self.get(position.get())),
            BodyTableSelector::Name(name) => {
                let mut matches = self.values.iter().filter(|table| table.name() == name);
                let Some(first) = matches.next() else {
                    return Ok(None);
                };
                if matches.next().is_some() {
                    return Err(BodyTableCatalogError::AmbiguousTableName);
                }
                Ok(Some(first))
            },
        }
    }

    /// Iterate over semantic table snapshots in rooted source order.
    pub fn iter(&self) -> std::slice::Iter<'_, BodyTableSnapshot> {
        self.values.iter()
    }
}

impl<'a> IntoIterator for &'a BodyTableCatalog {
    type Item = &'a BodyTableSnapshot;
    type IntoIter = std::slice::Iter<'a, BodyTableSnapshot>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

impl Package {
    /// Read the rooted Pages body-table catalog in source order.
    ///
    /// Discovery performs one bounded ownership walk and one shared wire
    /// budget across the complete catalog.  A valid body-less or table-less
    /// root returns an empty catalog; malformed rooted metadata fails closed.
    ///
    /// ```no_run
    /// use litchi_pages::Package;
    ///
    /// # fn read(source: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::from_bytes(source)?;
    /// let tables = package.body_tables()?;
    /// for table in &tables {
    ///     println!("{}: {} rows by {} columns", table.name(), table.rows(), table.columns());
    ///     let name = package.body_table_name(table.selector())?;
    ///     assert_eq!(name.as_str(), table.name());
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// The catalog is a read-only discovery snapshot.  A successful catalog
    /// read does not by itself admit every later mutation on that table.
    pub fn body_tables(&self) -> Result<BodyTableCatalog, BodyTableCatalogError> {
        let mut budget =
            table_lock::WireBudget::new(self.state.source.limits()).map_err(map_lock_error)?;
        budget
            .charge_source_catalog(&self.state.source)
            .map_err(map_lock_error)?;
        let entries = table_lock::body_table_catalog_with_budget(self, &mut budget)
            .map_err(map_lock_error)?;
        BodyTableCatalog::from_targets(entries, &mut budget)
    }
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableCatalogError {
    match error {
        table_lock::BodyTableLockError::AmbiguousTableName => {
            BodyTableCatalogError::AmbiguousTableName
        },
        table_lock::BodyTableLockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableCatalogError::LimitExceeded {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::BodyTableLockError::Allocation { amount } => {
            BodyTableCatalogError::Allocation { amount }
        },
        table_lock::BodyTableLockError::TableNotFound
        | table_lock::BodyTableLockError::AmbiguousSelector
        | table_lock::BodyTableLockError::UnsupportedSource
        | table_lock::BodyTableLockError::InvalidSource
        | table_lock::BodyTableLockError::Verification
        | table_lock::BodyTableLockError::PatchConflict => BodyTableCatalogError::InvalidSource,
    }
}

const fn map_lock_limit(kind: table_lock::BodyTableLockLimitKind) -> BodyTableCatalogLimitKind {
    use table_lock::BodyTableLockLimitKind as Lock;
    match kind {
        Lock::InputBytes => BodyTableCatalogLimitKind::InputBytes,
        Lock::OutputBytes => BodyTableCatalogLimitKind::OutputBytes,
        Lock::Entries => BodyTableCatalogLimitKind::Entries,
        Lock::EntryBytes => BodyTableCatalogLimitKind::EntryBytes,
        Lock::TotalEntryBytes => BodyTableCatalogLimitKind::TotalEntryBytes,
        Lock::PackageBytes => BodyTableCatalogLimitKind::PackageBytes,
        Lock::PayloadBytes => BodyTableCatalogLimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => BodyTableCatalogLimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => BodyTableCatalogLimitKind::PayloadObjects,
        Lock::PayloadMessages => BodyTableCatalogLimitKind::PayloadMessages,
        Lock::PayloadItems => BodyTableCatalogLimitKind::PayloadItems,
        Lock::PayloadReferences => BodyTableCatalogLimitKind::PayloadReferences,
        Lock::WireBytes => BodyTableCatalogLimitKind::WireBytes,
        Lock::WireFields => BodyTableCatalogLimitKind::WireFields,
        Lock::WireNesting => BodyTableCatalogLimitKind::WireNesting,
        Lock::WireWork => BodyTableCatalogLimitKind::WireWork,
    }
}
