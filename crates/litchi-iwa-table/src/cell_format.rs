//! Cell data formats and controls of one resolved table cell.
//!
//! The eleven display families (Number, Currency, Percentage, Scientific,
//! Fraction, Date and Time, Duration, Text, Custom, Numeral System and the
//! control display formats), the four scalar controls (Checkbox, Star
//! Rating, Slider, Stepper) and the Pop-Up Menu graph. Each read or rewrite
//! takes a [`TableSource`](crate::TableSource), a resolved [`Target`](crate::cell_format::Target), a
//! content-free [`Path`](crate::cell_format::Path) the leaf chooses, and one operation-local
//! [`TransactionBudget`](crate::cell_format::TransactionBudget). Rewrites return private archive bytes per changed
//! member; the leaf owns reassembly, candidate reopen and its public
//! transaction types.
//!
//! Moved from `litchi-numbers` (`package/table_cell_*_native.rs` and the
//! Pop-Up Menu engine) in record 0379. Numbers behaviour is unchanged.

use std::fmt;

use litchi_iwa_common::table::coordinate::At;
use thiserror::Error as ThisError;

mod budget;
/// Scalar-control graph surgery and the shared format-list primitives.
pub mod control;
/// The display-format families of an existing cell.
pub mod display;
mod metadata;
/// The Pop-Up Menu reads, native graph rewrite and metadata transition.
pub mod pop_up_menu;
/// The package metadata registry facts and transitions a Pop-Up Menu edit
/// needs.
pub mod pop_up_menu_metadata;
/// Native Pop-Up Menu graph transitions below the Pop-Up Menu engine.
pub mod pop_up_menu_native;

pub use budget::TransactionBudget;
pub use metadata::{ENTRY_NAME as METADATA_ENTRY_NAME, MessageRoute};

/// A content-free location carried by a table-cell format failure.
///
/// The leaf chooses the coordinates: `scope` is the position of the table's
/// container (a Numbers sheet, a Keynote slide; zero for a Pages body), and
/// `table` the table's position inside it. The engines never interpret them;
/// they only return them in errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Path {
    /// The complete package.
    Package,
    /// One selected table cell.
    Cell {
        /// Leaf-defined container position.
        scope: usize,
        /// Leaf-defined table position inside the container.
        table: usize,
        /// Zero-based cell coordinate.
        position: At,
    },
}

/// One resolved table cell: the table model's native route and lock state.
#[derive(Debug, Clone, Copy)]
pub struct Target {
    /// Leaf-defined container position, returned in [`Path::Cell`].
    pub scope: usize,
    /// Leaf-defined table position, returned in [`Path::Cell`].
    pub table: usize,
    /// Zero-based cell coordinate.
    pub position: At,
    /// Object identifier of the `TST.TableModelArchive`.
    pub model_identifier: u64,
    /// Component of the table model in [`TableSource::catalog`](crate::TableSource::catalog).
    pub component_index: usize,
    /// Object index of the table model inside its component.
    pub object_index: usize,
    /// Message index of the table model inside its object.
    pub message_index: usize,
    /// Native message type of the table model.
    pub message_type: u32,
    /// Whether the table is effectively locked.
    pub locked: bool,
}

impl Target {
    /// The content-free location of this cell.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Cell {
            scope: self.scope,
            table: self.table,
            position: self.position,
        }
    }
}

/// A finite resource governed by a table-cell format transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete candidate package output bytes.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// Bytes in one physical entry.
    EntryBytes,
    /// Aggregate physical entry bytes.
    TotalEntryBytes,
    /// Physical package/container metadata bytes.
    PackageBytes,
    /// Decoded native payload bytes.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native framing/items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Strict codec input bytes.
    WireBytes,
    /// Strict codec output bytes.
    WireOutputBytes,
    /// Strict codec reference-envelope bytes.
    WireReferenceBytes,
    /// Strict codec selected-text bytes.
    WireTextBytes,
    /// Strict codec fields inspected.
    WireFields,
    /// Strict codec nesting depth.
    WireNesting,
    /// Strict codec work.
    WireWork,
    /// Private codec/reassembly scratch bytes.
    ScratchBytes,
    /// Private candidate bytes retained during execution.
    RetainedBytes,
    /// Bounded allocation units reserved by a phase.
    Allocations,
    /// Compressed physical member bytes.
    CompressedBytes,
    /// Aggregate transaction work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A content-redacted table-cell format failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
pub enum Error {
    /// A well-formed request the engine refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// Rooted ownership or wire framing is invalid.
    #[error("the table-cell format source is invalid at {path:?}")]
    InvalidSource {
        /// Content-free failure location.
        path: Path,
    },
    /// A finite transaction resource ceiling was exceeded.
    #[error("table-cell format {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
        /// Content-free failure location.
        path: Path,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the table-cell format transaction")]
    Allocation {
        /// Requested units.
        amount: usize,
        /// Content-free failure location.
        path: Path,
    },
    /// Candidate verification failed.
    #[error("the edited table-cell format failed verification")]
    Verification,
    /// A patch or staged value does not match the exact source.
    #[error("the table-cell format edit does not match the exact source")]
    PatchConflict,
}

/// A well-formed request a table-cell format engine refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
pub enum Refusal {
    /// The selected cell was not present in the table.
    #[error("the selected table cell was not found")]
    CellNotFound,
    /// A changed operation targeted an effectively locked table.
    #[error("the selected table is locked at {path:?}")]
    TableLocked {
        /// Content-free failure location.
        path: Path,
    },
    /// A native dependency is not safe to rewrite.
    #[error("the selected table cell has an unsupported dependency at {path:?}")]
    UnsupportedDependency {
        /// Content-free failure location.
        path: Path,
    },
    /// The exact native owner is not available for this source.
    #[error("this source does not support exact table-cell format editing")]
    UnsupportedSource,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// Preserve resource and allocation failures from the bounded Snappy
/// encoder at the table transaction boundary.  Compression uses the same
/// retained package limits as decoding, so a caller-selected compressed
/// entry ceiling is an ordinary typed refusal rather than a verification
/// failure.
pub fn map_core_error(error: litchi_iwa_core::Error, path: Path) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_core_limit_kind(kind),
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => Error::Allocation {
            amount: requested,
            path,
        },
        _ => Error::InvalidSource { path },
    }
}

pub const fn map_core_limit_kind(kind: litchi_iwa_core::LimitKind) -> LimitKind {
    match kind {
        litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
        litchi_iwa_core::LimitKind::Messages | litchi_iwa_core::LimitKind::MessagesPerObject => {
            LimitKind::PayloadMessages
        },
        litchi_iwa_core::LimitKind::HeaderFields
        | litchi_iwa_core::LimitKind::MetadataItems
        | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::PayloadItems,
        litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
        litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
        | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => LimitKind::CompressedBytes,
        litchi_iwa_core::LimitKind::ArchiveBytes
        | litchi_iwa_core::LimitKind::ObjectBytes
        | litchi_iwa_core::LimitKind::MessageBytes
        | litchi_iwa_core::LimitKind::HeaderBytes
        | litchi_iwa_core::LimitKind::HeaderMemoryBytes
        | litchi_iwa_core::LimitKind::SnappyChunkBytes
        | litchi_iwa_core::LimitKind::SnappyStreamBytes => LimitKind::PayloadBytes,
    }
}

/// Derive the BNC cell profile from the owner-selected source extent.
///
/// Fixed-width field checks are derived from the retained source extent; the
/// output ceiling stays at the BNC codec's hard maximum because an owner may
/// add metadata fields during a rewrite.
pub(crate) fn bnc_limits_for_source(source: &[u8]) -> litchi_numbers_wire::Limits {
    if source.is_empty() {
        return litchi_numbers_wire::Limits::DEFAULT;
    }
    litchi_numbers_wire::Limits::DEFAULT
        .with_input_bytes(source.len())
        .and_then(|limits| limits.with_tail_bytes(source.len()))
        .expect("source-derived BNC limits stay within the checked hard ceilings")
}

const MIN_SIGN_EXTENDED_I32: u64 = u64::MAX - 2_147_483_647;

fn canonical_varint(source: &[u8]) -> Result<u64, ()> {
    let (value, length) = litchi_iwa_common::decode_varint_from_bytes(source).map_err(|_| ())?;
    if length != source.len() || litchi_iwa_common::varint::encoded_len(value) != length {
        return Err(());
    }
    Ok(value)
}

/// Decode one canonical local `TSP.Reference` payload to its identifier.
pub(crate) fn local_reference_identifier(source: &[u8]) -> Result<u64, ()> {
    let view =
        litchi_iwa_common::wire::WireView::parse(source, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|_| ())?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(|_| ())?;
        match field.number() {
            1 if identifier.is_none() && field.wire_type() == 0 => {
                identifier = Some(canonical_varint(field.payload())?);
            },
            2 if deprecated_type.is_none() && field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > u64::from(i32::MAX.unsigned_abs()) && value < MIN_SIGN_EXTENDED_I32 {
                    return Err(());
                }
                deprecated_type = Some(value);
            },
            3 if external.is_none() && field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > 1 {
                    return Err(());
                }
                external = Some(value != 0);
            },
            1..=3 => return Err(()),
            _ => {},
        }
    }
    let resolved = identifier.ok_or(())?;
    if resolved == 0 || external == Some(true) {
        return Err(());
    }
    Ok(resolved)
}

/// The exact physical source, or the typed refusal when the package was not
/// opened from bytes.
pub(crate) fn physical_source(
    source: &dyn crate::TableSource,
) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
    source
        .physical()
        .ok_or(Error::Refused(Refusal::UnsupportedSource))
}

/// A package opened from bytes for this crate's unit tests: a linear object
/// lookup over the catalog and the default ingress limits.
#[cfg(test)]
pub(crate) mod test_source {
    use litchi_iwa_archive::{ComponentCatalog, Limits, SourceCatalog};

    use crate::{ResolveError, ResolvedObject, TableSource};

    pub(crate) struct Opened {
        catalog: SourceCatalog,
        bytes: Vec<u8>,
    }

    impl Opened {
        pub(crate) fn open(path: &std::path::Path) -> Self {
            let bytes = std::fs::read(path).expect("fixture bytes");
            let catalog = SourceCatalog::from_bytes(&bytes, Limits::default()).expect("catalog");
            Self { catalog, bytes }
        }
    }

    impl TableSource for Opened {
        fn catalog(&self) -> &ComponentCatalog {
            self.catalog.components()
        }

        fn physical(&self) -> Option<&SourceCatalog> {
            Some(&self.catalog)
        }

        fn resolve_object(
            &self,
            identifier: u64,
        ) -> Result<Option<ResolvedObject<'_>>, ResolveError> {
            for (component_index, component) in self.catalog().iter().enumerate() {
                for (object_index, object) in component.archive().objects.iter().enumerate() {
                    if object.archive_info.identifier == Some(identifier) {
                        return Ok(Some(ResolvedObject {
                            messages: &object.messages,
                            component_index,
                            object_index,
                        }));
                    }
                }
            }
            Ok(None)
        }

        fn source_bytes(&self) -> &[u8] {
            &self.bytes
        }

        fn archive_limits(&self) -> Limits {
            Limits::default()
        }

        fn max_objects(&self) -> usize {
            1_000_000
        }

        fn max_references(&self) -> usize {
            1_000_000
        }
    }
}
