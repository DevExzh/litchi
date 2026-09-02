//! Archive-free physical row-sort values for a Keynote slide table.
//!
//! The physical sorter is deliberately a separate semantic surface from the
//! persisted [`crate::slide::table::sort`] configuration. A physical
//! operation moves body rows and every admitted row-affine structure as one
//! exact-source transaction. Native object identifiers, package members,
//! protobuf messages, and encoded bytes stay below the package adapter
//! boundary.
//!
//! The package operation consumes the persisted
//! [`crate::slide::table::sort::Order`] rather than inventing a second sort
//! configuration. The current owner admits only
//! existing package-backed tables whose addressed sort keys are present and
//! scalar. For each rule, all selected rows must use one scalar kind from
//! text, finite number, boolean, date, or duration; different rules may use
//! different kinds. Text comparison is Rust's ordinal string ordering,
//! boolean comparison is `false < true`, and numeric/date/duration comparison
//! is deterministic `f64::total_cmp` ordering. Formulas, errors, rich text,
//! unsupported values, missing keys, mixed kinds within a rule, and native
//! row-affine structures outside the admitted topology are refused before
//! publication. This is a bounded physical transaction, not a general
//! Keynote table editor or a native-application compatibility guarantee.

use std::fmt;

use litchi_core::Position;

/// Failures while constructing the shared checked sort values.
pub use super::sort::Error as SortValueError;
pub use super::sort::{ColumnIndex, Direction, Order, RowRange, Rule, Scope};

/// A content-free semantic location associated with a physical table sort.
///
/// `Rows` is body-relative and half-open.  Header and footer rows are never
/// represented by this path.  The range itself is checked for non-emptiness
/// by [`RowRange::new`]; the package adapter validates it against the selected
/// table's body length before planning. Table positions are ordinal among
/// table drawables in the slide's z-order; non-table drawables do not consume
/// a table position.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Keynote package.
    Package,
    /// One selected slide/table root.
    Table {
        /// Zero-based slide position.
        slide: Position,
        /// Zero-based table position among the slide's z-order tables.
        table: Position,
    },
    /// A body-relative row range within one selected slide/table root.
    Rows {
        /// Zero-based slide position.
        slide: Position,
        /// Zero-based table position among the slide's z-order tables.
        table: Position,
        /// Non-empty body-relative half-open row range.
        range: RowRange,
    },
}

impl Path {
    /// Construct the package-level path.
    #[must_use]
    pub const fn package() -> Self {
        Self::Package
    }

    /// Construct a path for one selected table.
    #[must_use]
    pub const fn table(slide: Position, table: Position) -> Self {
        Self::Table { slide, table }
    }

    /// Construct a path for a selected body-row range.
    #[must_use]
    pub const fn rows(slide: Position, table: Position, range: RowRange) -> Self {
        Self::Rows {
            slide,
            table,
            range,
        }
    }

    /// Return the checked slide position when this path addresses a table.
    #[must_use]
    pub const fn slide(self) -> Option<Position> {
        match self {
            Self::Package => None,
            Self::Table { slide, .. } | Self::Rows { slide, .. } => Some(slide),
        }
    }

    /// Return the checked table position when this path addresses a table.
    #[must_use]
    pub const fn table_position(self) -> Option<Position> {
        match self {
            Self::Package => None,
            Self::Table { table, .. } | Self::Rows { table, .. } => Some(table),
        }
    }

    /// Return the selected row range, if this path is row-scoped.
    #[must_use]
    pub const fn range(self) -> Option<RowRange> {
        match self {
            Self::Package | Self::Table { .. } => None,
            Self::Rows { range, .. } => Some(range),
        }
    }
}

impl fmt::Display for Path {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Table { slide, table } => {
                write!(formatter, "slide {} table {}", slide.get(), table.get())
            },
            Self::Rows {
                slide,
                table,
                range,
            } => write!(
                formatter,
                "slide {} table {} rows {}..{}",
                slide.get(),
                table.get(),
                range.start(),
                range.end()
            ),
        }
    }
}

/// A row-affine or key-value feature that the current physical owner does
/// not rewrite completely.
///
/// This enum intentionally carries no authored text, native identifier, or
/// wire representation.  It lets an adapter report a stable, content-free
/// reason while retaining the option to admit more feature families later.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum UnsupportedFeature {
    /// A formula-bearing cell or formula dependency graph.
    Formula,
    /// A formula error or its dependency sidecar.
    FormulaError,
    /// A merged-cell region or merge owner.
    MergedCells,
    /// A filtered row state.
    FilteredRows,
    /// A grouped or outline row state.
    GroupedRows,
    /// A pivot-table graph.
    PivotTable,
    /// A category-grouping graph.
    CategoryGrouping,
    /// A spill or array-expansion graph.
    SpillRanges,
    /// A conditional-style rule or applied-rule graph.
    ConditionalStyles,
    /// A hidden-axis state that is not positional and UID-addressed.
    NonPositionalHiddenState,
    /// A row-affine dependency not covered by the current owner.
    RowAffineDependency,
    /// A comment anchor whose location is not carried safely with its cell.
    CommentAnchors,
    /// A rich-text or other non-scalar value used as a sort key.
    UnsupportedSortKey,
    /// A body cell value outside the admitted scalar domains.
    UnsupportedCell,
    /// A table topology that cannot be proven lossless for row movement.
    UnsupportedTopology,
}

impl fmt::Display for UnsupportedFeature {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Formula => "formula cells",
            Self::FormulaError => "formula errors",
            Self::MergedCells => "merged cells",
            Self::FilteredRows => "filtered rows",
            Self::GroupedRows => "grouped rows",
            Self::PivotTable => "pivot tables",
            Self::CategoryGrouping => "category grouping",
            Self::SpillRanges => "spill ranges",
            Self::ConditionalStyles => "conditional styles",
            Self::NonPositionalHiddenState => "non-positional hidden state",
            Self::RowAffineDependency => "row-affine dependency",
            Self::CommentAnchors => "comment anchors",
            Self::UnsupportedSortKey => "unsupported sort key",
            Self::UnsupportedCell => "unsupported cell value",
            Self::UnsupportedTopology => "unsupported table topology",
        })
    }
}

/// Finite resources governed by one physical slide-table sort transaction.
///
/// The concrete package adapter maps its internal archive and wire budgets to
/// these stable semantic categories.  The categories are intentionally
/// broader than one implementation so a future Buffa owner can preserve the
/// same safety contract without changing the public error shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source package bytes inspected.
    InputBytes,
    /// Complete candidate package bytes produced.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// Bytes in one physical entry.
    EntryBytes,
    /// Aggregate physical-entry bytes.
    TotalEntryBytes,
    /// Package container and metadata bytes.
    PackageBytes,
    /// Bytes in decoded payload containers.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Decoded payload objects.
    PayloadObjects,
    /// Decoded payload messages.
    PayloadMessages,
    /// Decoded payload items or metadata records.
    PayloadItems,
    /// Native object-reference edges inspected.
    PayloadReferences,
    /// Protobuf/Buffa input bytes inspected.
    WireBytes,
    /// Protobuf/Buffa output bytes emitted.
    WireOutputBytes,
    /// Wire fields inspected or emitted.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Wire traversal or rewrite work.
    WireWork,
    /// Number of sort rules admitted.
    WireRules,
    /// Number of table columns admitted.
    WireColumns,
    /// Number of fallible wire allocations admitted.
    WireAllocations,
    /// Bytes retained by a lazy projection or transaction.
    WireRetainedBytes,
    /// Temporary wire scratch bytes.
    WireScratchBytes,
    /// Aggregate physical transaction work.
    TransactionWork,
    /// Body rows admitted to the physical operation.
    PhysicalRows,
    /// Table columns admitted to the physical operation.
    PhysicalColumns,
    /// Sort-key cells inspected.
    KeyCells,
    /// BNC offsets inspected or rewritten.
    BncOffsets,
    /// Rewritten package components.
    Components,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::WireRules => "wire rules",
            Self::WireColumns => "wire columns",
            Self::WireAllocations => "wire allocations",
            Self::WireRetainedBytes => "wire retained bytes",
            Self::WireScratchBytes => "wire scratch bytes",
            Self::TransactionWork => "transaction work",
            Self::PhysicalRows => "physical rows",
            Self::PhysicalColumns => "physical columns",
            Self::KeyCells => "sort-key cells",
            Self::BncOffsets => "BNC offsets",
            Self::Components => "components",
        })
    }
}

/// Failure from a physical Keynote slide-table sort read or transaction.
///
/// Every contextual field is a checked semantic value.  In particular, no
/// variant stores a slide name, native ID, archive member, protobuf message,
/// or authored cell content; `Debug` and `Display` are therefore safe for
/// diagnostics and logs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The source does not have exact physical provenance for this operation.
    #[error("this Keynote source does not support physical slide-table sorting")]
    UnsupportedSource,
    /// A known feature cannot be rewritten completely by this owner.
    #[error("unsupported Keynote slide-table sort feature {feature} at {path}")]
    UnsupportedFeature {
        /// Content-free semantic location.
        path: Path,
        /// Rejected row-affine or key feature.
        feature: UnsupportedFeature,
    },
    /// A row-affine dependency is valid but outside the owner's rewrite
    /// contract.
    #[error("unsupported Keynote slide-table sort dependency {feature} at {path}")]
    UnsupportedDependency {
        /// Content-free semantic location.
        path: Path,
        /// Rejected row-affine dependency.
        feature: UnsupportedFeature,
    },
    /// The rooted physical table topology is not admitted.
    #[error("unsupported Keynote slide-table sort topology at {path}")]
    UnsupportedTopology {
        /// Content-free semantic location.
        path: Path,
    },
    /// A semantic selector resolved to more than one physical table.
    #[error("the Keynote slide-table sort selector is ambiguous")]
    AmbiguousSelector,
    /// A slide-name selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide matched a semantic name selector.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked slide position was outside the presentation.
    #[error("the Keynote show has no slide at position {}", position.get())]
    SlidePositionNotFound {
        /// Requested zero-based slide position.
        position: Position,
    },
    /// A checked table position was outside the selected slide.
    #[error("the selected Keynote slide has no table at position {}", position.get())]
    TablePositionNotFound {
        /// Requested zero-based table position.
        position: Position,
    },
    /// A changed operation targeted a locked table.
    #[error("the selected Keynote slide table is locked at {path}")]
    TableLocked {
        /// Content-free semantic location.
        path: Path,
    },
    /// No persisted sort order was available for the physical operation.
    #[error("the selected Keynote slide table has no persisted sort order at {path}")]
    SortOrderMissing {
        /// Content-free semantic location.
        path: Path,
    },
    /// The configured order scope does not match the physical request.
    #[error(
        "Keynote slide-table sort scope mismatch at {path}: configured {configured:?}, requested {requested:?}"
    )]
    ScopeMismatch {
        /// Content-free semantic location.
        path: Path,
        /// Scope encoded by the persisted order.
        configured: Scope,
        /// Scope requested by the physical operation.
        requested: Scope,
    },
    /// A body-relative row range exceeds the selected table body.
    #[error("Keynote slide-table sort row range {range:?} exceeds {body_rows} body rows at {path}")]
    RowRangeOutOfBounds {
        /// Content-free semantic location.
        path: Path,
        /// Checked, non-empty body-relative range.
        range: RowRange,
        /// Declared body-row count.
        body_rows: usize,
    },
    /// The source graph, dimensions, references, or wire framing is invalid.
    #[error("the Keynote slide-table sort source is invalid at {path}")]
    InvalidSource {
        /// Content-free semantic location.
        path: Path,
    },
    /// A sort key is missing from one body row.
    #[error("the Keynote slide-table sort key is missing at {path}")]
    MissingSortKey {
        /// Content-free semantic location of the selected table/range.
        path: Path,
        /// Body-relative row containing the missing key.
        row: Position,
        /// Zero-based physical key column.
        column: ColumnIndex,
    },
    /// A sort key is outside the scalar domains supported by this owner.
    #[error("the Keynote slide-table sort key is unsupported at {path}")]
    UnsupportedCell {
        /// Content-free semantic location of the selected table/range.
        path: Path,
        /// Body-relative row containing the unsupported key.
        row: Position,
        /// Zero-based physical key column.
        column: ColumnIndex,
    },
    /// One configured rule sees incompatible scalar domains across rows.
    #[error("the Keynote slide-table sort key domains disagree at {path}")]
    MixedSortKeyKinds {
        /// Content-free semantic location of the selected table/range.
        path: Path,
        /// Body-relative row containing the mismatched key.
        row: Position,
        /// Zero-based rule position.
        rule: Position,
    },
    /// The planner did not produce a complete row bijection.
    #[error("the Keynote slide-table sort produced an invalid row permutation at {path}")]
    InvalidPermutation {
        /// Content-free semantic location.
        path: Path,
    },
    /// Candidate reopening or semantic locality verification failed.
    #[error("the edited Keynote slide-table sort failed semantic verification at {path}")]
    Verification {
        /// Content-free semantic location.
        path: Path,
    },
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote slide-table sort {kind} limit exceeded at {path}: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
        /// Content-free semantic location.
        path: Path,
    },
    /// A bounded allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote slide-table sort at {path}")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
        /// Content-free semantic location.
        path: Path,
    },
    /// A patch was applied to a package other than its exact source.
    #[error("the Keynote slide-table sort patch does not match the exact source package")]
    PatchConflict,
}

/// Content-free observations from one physical row-sort publication.
///
/// The counters describe the admitted transaction, not the size of any
/// native identifier graph.  An unchanged operation reports zero movement,
/// zero touched components, no preview deletion, and no full reopen.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    moved_rows: usize,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    /// Return whether the published package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of body-row envelopes whose destination changed.
    #[must_use]
    pub const fn moved_rows(self) -> usize {
        self.moved_rows
    }

    /// Return the number of body-row envelopes whose destination changed.
    ///
    /// This spelling mirrors the wording used by physical-sort callers;
    /// [`Self::moved_rows`] remains the canonical accessor.
    #[must_use]
    pub const fn rows_moved(self) -> usize {
        self.moved_rows()
    }

    /// Return the number of admitted physical components participating in
    /// the verified rewrite closure.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of canonical previews deleted for this operation.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the candidate was fully reopened and semantically
    /// reselected before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }

    /// Construct diagnostics for an exact no-op.
    pub(crate) const fn unchanged() -> Self {
        Self {
            changed: false,
            moved_rows: 0,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    /// Construct diagnostics for a verified physical publication.
    pub(crate) const fn published(
        moved_rows: usize,
        touched_components: usize,
        deleted_previews: usize,
    ) -> Self {
        Self {
            changed: true,
            moved_rows,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }
}

/// Result type for a physical Keynote slide-table sort transaction.
pub type Result<T> = std::result::Result<T, Error>;

/// Compatibility spelling used by callers that distinguish value errors from
/// package transaction errors in generic code.
pub type TransactionError = Error;

/// Exact-source physical row-sort transaction types.
///
/// The package adapter owns the package-bearing `Edit`, `Patch`, and `Commit`
/// values.  Their public aliases are kept here so callers can stay in the
/// semantic `slide::table::physical_sort` namespace while the implementation
/// remains below the archive boundary.
pub mod transaction {
    pub use super::{Diagnostics, Error, LimitKind, Path, TransactionError, UnsupportedFeature};
    pub use crate::{
        SlideTablePhysicalSortCommit as Commit, SlideTablePhysicalSortEdit as Edit,
        SlideTablePhysicalSortPatch as Patch,
    };
}
