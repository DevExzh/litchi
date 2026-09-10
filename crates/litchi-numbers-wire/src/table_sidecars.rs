//! Shared, bounded readers for selected table-data-list sidecars.
//!
//! A format adapter owns the package graph and the object lookup that supplies
//! borrowed messages.  This module owns the part that is identical for
//! Numbers, Pages, and Keynote: strict root/segment list selection, entry
//! shape validation, bounded retention of sidecar values, and conversion of a
//! classified cell into the archive-free table vocabulary.  Native object
//! identifiers are kept behind [`SidecarReference`] and never appear in the
//! common table read returned by an adapter.

use std::collections::HashSet;
use std::fmt;
use std::num::NonZeroU64;

use litchi_iwa_common::table::cell::value::{FiniteF64, Value};
use litchi_iwa_common::table::read::{CellComment, Comment, CommentTimestamp};
use litchi_iwa_protos::comment_storage_codec;
use litchi_iwa_protos::numbers_formula_codec;
use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;

use crate::cell_value::{CellValueSource, ValueSource};
use crate::formula_envelope;
use crate::formula_render::{self, FormulaEventRenderBudget, ReferenceResolver};
use crate::table_data_list::{self, Message, RootOrSegment};

/// Native type number for `TSD.CommentStorageArchive` messages.
pub const COMMENT_STORAGE_MESSAGE_KIND: u32 = 3_056;

/// A selected table-data-list sidecar.
#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub enum SidecarKind {
    /// The shared string table.
    Strings,
    /// Formula archives, retained as source bytes until a cell needs them.
    Formulas,
    /// Formula error display strings.
    FormulaErrors,
    /// Rich-text payload references.
    RichTextPayloads,
    /// Cell comment-storage references.
    Comments,
}

impl SidecarKind {
    /// Return the native `TableDataList.ListType` value.
    #[must_use]
    pub const fn list_type(self) -> i32 {
        match self {
            Self::Strings => 1,
            Self::Formulas => 3,
            Self::FormulaErrors => 5,
            Self::RichTextPayloads => 8,
            Self::Comments => 10,
        }
    }
}

impl fmt::Display for SidecarKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Strings => "strings",
            Self::Formulas => "formulas",
            Self::FormulaErrors => "formula errors",
            Self::RichTextPayloads => "rich text payloads",
            Self::Comments => "comments",
        })
    }
}

/// A validated native object reference retained only for adapter-side lazy
/// payload lookup.
///
/// This type is intentionally opaque in diagnostics and is not part of the
/// archive-free table read.  Format adapters may pass it back to their own
/// object resolver when resolving rich text, comments, or replies.
#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct SidecarReference(NonZeroU64);

impl SidecarReference {
    /// Construct a reference, rejecting the native zero absence sentinel.
    #[must_use]
    pub const fn new(identifier: u64) -> Option<Self> {
        match NonZeroU64::new(identifier) {
            Some(identifier) => Some(Self(identifier)),
            None => None,
        }
    }

    /// Return the native identifier at the package-adapter boundary.
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.0.get()
    }
}

/// One retained sidecar value.
///
/// Text and formula bytes are copied into bounded owned buffers because the
/// generated-free visitor callback is allowed to borrow for only the callback
/// invocation.  Formula bytes remain unparsed until a cell asks for them.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum SidecarValue {
    /// A UTF-8 string-table or formula-error value.
    Text(Box<str>),
    /// A formula archive retained as exact source bytes.
    Formula(Box<[u8]>),
    /// A reference to a package-owned rich-text or comment-storage payload.
    Reference(SidecarReference),
}

impl SidecarValue {
    /// Return retained text when this is a text sidecar value.
    #[must_use]
    pub fn text(&self) -> Option<&str> {
        match self {
            Self::Text(value) => Some(value),
            Self::Formula(_) | Self::Reference(_) => None,
        }
    }

    /// Return retained formula bytes when this is a formula sidecar value.
    #[must_use]
    pub fn formula(&self) -> Option<&[u8]> {
        match self {
            Self::Formula(value) => Some(value),
            Self::Text(_) | Self::Reference(_) => None,
        }
    }

    /// Return the payload reference when this is a reference sidecar value.
    #[must_use]
    pub const fn reference(&self) -> Option<SidecarReference> {
        match self {
            Self::Reference(value) => Some(*value),
            Self::Text(_) | Self::Formula(_) => None,
        }
    }
}

/// A sorted, immutable sidecar list.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct SidecarList {
    kind: SidecarKind,
    entries: Box<[(u32, SidecarValue)]>,
}

impl SidecarList {
    /// Construct a list from coordinator-sorted entries.
    ///
    /// The list is checked for duplicate keys before publication.  The
    /// caller should normally use [`read_sidecar_list`] so the strict list
    /// coordinator performs that check together with segment validation.
    pub fn from_entries(
        kind: SidecarKind,
        mut entries: Vec<(u32, SidecarValue)>,
    ) -> Result<Self, SidecarIssue> {
        if entries
            .iter()
            .any(|(_, value)| !value_matches_kind(value, kind))
        {
            return Err(SidecarIssue::InvalidEntry { kind });
        }
        entries.sort_unstable_by_key(|(key, _)| *key);
        if entries.windows(2).any(|pair| pair[0].0 == pair[1].0) {
            return Err(SidecarIssue::DuplicateKey { kind });
        }
        Ok(Self {
            kind,
            entries: entries.into_boxed_slice(),
        })
    }

    /// Return the list's sidecar kind.
    #[must_use]
    pub const fn kind(&self) -> SidecarKind {
        self.kind
    }

    /// Return the number of retained entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Return whether no entries were retained.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Look up one sidecar entry by its cell-local key.
    #[must_use]
    pub fn get(&self, key: u32) -> Option<&SidecarValue> {
        self.entries
            .binary_search_by_key(&key, |(entry_key, _)| *entry_key)
            .ok()
            .map(|index| &self.entries[index].1)
    }

    /// Iterate sorted sidecar keys and values.
    #[must_use]
    pub fn iter(&self) -> impl ExactSizeIterator<Item = (u32, &SidecarValue)> + '_ {
        self.entries.iter().map(|(key, value)| (*key, value))
    }
}

fn value_matches_kind(value: &SidecarValue, kind: SidecarKind) -> bool {
    matches!(
        (kind, value),
        (
            SidecarKind::Strings | SidecarKind::FormulaErrors,
            SidecarValue::Text(_)
        ) | (SidecarKind::Formulas, SidecarValue::Formula(_))
            | (
                SidecarKind::RichTextPayloads | SidecarKind::Comments,
                SidecarValue::Reference(_)
            )
    )
}

/// The five selected sidecar maps needed by a semantic table projection.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct SidecarTables {
    strings: Option<SidecarList>,
    formulas: Option<SidecarList>,
    formula_errors: Option<SidecarList>,
    rich_text_payloads: Option<SidecarList>,
    comments: Option<SidecarList>,
}

impl SidecarTables {
    /// Construct empty sidecar maps.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            strings: None,
            formulas: None,
            formula_errors: None,
            rich_text_payloads: None,
            comments: None,
        }
    }

    /// Insert one list, rejecting a repeated sidecar kind.
    pub fn insert(&mut self, list: SidecarList) -> Result<(), SidecarIssue> {
        let slot = match list.kind() {
            SidecarKind::Strings => &mut self.strings,
            SidecarKind::Formulas => &mut self.formulas,
            SidecarKind::FormulaErrors => &mut self.formula_errors,
            SidecarKind::RichTextPayloads => &mut self.rich_text_payloads,
            SidecarKind::Comments => &mut self.comments,
        };
        if slot.is_some() {
            return Err(SidecarIssue::DuplicateList { kind: list.kind() });
        }
        *slot = Some(list);
        Ok(())
    }

    /// Return one selected list when it has been loaded.
    #[must_use]
    pub const fn list(&self, kind: SidecarKind) -> Option<&SidecarList> {
        match kind {
            SidecarKind::Strings => self.strings.as_ref(),
            SidecarKind::Formulas => self.formulas.as_ref(),
            SidecarKind::FormulaErrors => self.formula_errors.as_ref(),
            SidecarKind::RichTextPayloads => self.rich_text_payloads.as_ref(),
            SidecarKind::Comments => self.comments.as_ref(),
        }
    }

    /// Return the string sidecar entry for a cell key.
    #[must_use]
    pub fn string(&self, key: u32) -> Option<&str> {
        self.list(SidecarKind::Strings)
            .and_then(|list| list.get(key))
            .and_then(SidecarValue::text)
    }

    /// Return exact retained formula bytes for a cell key.
    #[must_use]
    pub fn formula(&self, key: u32) -> Option<&[u8]> {
        self.list(SidecarKind::Formulas)
            .and_then(|list| list.get(key))
            .and_then(SidecarValue::formula)
    }

    /// Return the formula-error text for a cell key.
    #[must_use]
    pub fn formula_error(&self, key: u32) -> Option<&str> {
        self.list(SidecarKind::FormulaErrors)
            .and_then(|list| list.get(key))
            .and_then(SidecarValue::text)
    }

    /// Return the rich-text payload object for a cell key.
    #[must_use]
    pub fn rich_text_reference(&self, key: u32) -> Option<SidecarReference> {
        self.list(SidecarKind::RichTextPayloads)
            .and_then(|list| list.get(key))
            .and_then(SidecarValue::reference)
    }

    /// Return the comment-storage object for a cell key.
    #[must_use]
    pub fn comment_reference(&self, key: u32) -> Option<SidecarReference> {
        self.list(SidecarKind::Comments)
            .and_then(|list| list.get(key))
            .and_then(SidecarValue::reference)
    }
}

/// Allocation categories debited by a sidecar reader before a fallible
/// staging operation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SidecarAllocation {
    /// A sidecar key set entry.
    Keys,
    /// A retained `(key, value)` pair.
    Values,
    /// A root-referenced segment identity.
    Segments,
    /// Copied UTF-8 sidecar text.
    Text,
    /// Copied formula source bytes.
    FormulaBytes,
    /// Comment or reply records.
    Comments,
}

/// Neutral failures that an adapter maps into its own error vocabulary.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum SidecarIssue {
    /// A strict table-data-list codec pass failed.
    StorageDecode {
        /// Sidecar being decoded.
        kind: SidecarKind,
        /// Strict codec failure.
        error: storage::DecodeError,
    },
    /// A strict comment-storage codec pass failed.
    CommentDecode(comment_storage_codec::DecodeError),
    /// A strict formula codec pass failed.
    FormulaDecode(numbers_formula_codec::DecodeError),
    /// A shared root/segment coordinator rejected the selected topology.
    Coordinator(table_data_list::CoordinatorIssue),
    /// One list entry has no payload matching its selected list kind.
    InvalidEntry { kind: SidecarKind },
    /// A required object reference contains the native zero sentinel.
    ZeroReference { kind: SidecarKind },
    /// A present storage UUID contains two zero words.
    ZeroUuid,
    /// The selected object has no canonical payload of the required kind.
    MissingPayload { kind: SidecarKind },
    /// A selected object has more than one canonical payload.
    DuplicatePayload { kind: SidecarKind },
    /// Two entries in one sidecar use the same key.
    DuplicateKey { kind: SidecarKind },
    /// A root repeats one segment object identity.
    DuplicateSegmentReference { kind: SidecarKind },
    /// Two sidecar lists with the same semantic role were inserted.
    DuplicateList { kind: SidecarKind },
    /// A bounded staging allocation failed.
    Allocation {
        /// Sidecar operation that failed.
        kind: SidecarKind,
        /// Collection or buffer that could not grow.
        target: SidecarAllocation,
        /// Requested units.
        amount: usize,
    },
    /// A comment reply list was truncated or contained an invalid identity.
    InvalidReplies,
}

impl fmt::Display for SidecarIssue {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::StorageDecode { kind, .. } => write!(formatter, "invalid {kind} list"),
            Self::CommentDecode(_) => formatter.write_str("invalid comment-storage payload"),
            Self::FormulaDecode(_) => formatter.write_str("invalid formula payload"),
            Self::Coordinator(_) => formatter.write_str("invalid sidecar list topology"),
            Self::InvalidEntry { kind } => write!(formatter, "invalid {kind} list entry"),
            Self::ZeroReference { kind } => write!(formatter, "zero reference in {kind} list"),
            Self::ZeroUuid => formatter.write_str("zero UUID in comment storage"),
            Self::MissingPayload { kind } => write!(formatter, "missing {kind} payload"),
            Self::DuplicatePayload { kind } => write!(formatter, "duplicate {kind} payload"),
            Self::DuplicateKey { kind } => write!(formatter, "duplicate key in {kind} list"),
            Self::DuplicateSegmentReference { kind } => {
                write!(formatter, "duplicate segment reference in {kind} list")
            },
            Self::DuplicateList { kind } => write!(formatter, "duplicate {kind} list"),
            Self::Allocation {
                kind,
                target,
                amount,
            } => write!(
                formatter,
                "could not allocate {amount} units for {kind} {target:?}"
            ),
            Self::InvalidReplies => formatter.write_str("invalid comment replies"),
        }
    }
}

impl std::error::Error for SidecarIssue {}

/// Budget and error hooks supplied by one concrete format adapter.
///
/// The same mutable implementation is passed through every list and payload
/// read.  Implementations should update their cumulative ledger only after a
/// strict report is available, while retaining attempted work on failures.
pub trait SidecarReadBudget {
    /// Adapter error type.
    type Error;

    /// Build strict table-data-list options from the current residual ledger.
    fn list_options(
        &mut self,
        kind: SidecarKind,
        source: &[u8],
    ) -> Result<storage::DecodeOptions, Self::Error>;

    /// Charge one completed table-data-list codec report.
    fn charge_list_report(
        &mut self,
        kind: SidecarKind,
        report: storage::DecodeReport,
    ) -> Result<(), Self::Error>;

    /// Charge a routing/probe pass without charging retained references or
    /// text.  A type probe must remain cheap even when it rejects a candidate.
    fn charge_list_probe_report(
        &mut self,
        kind: SidecarKind,
        report: storage::DecodeReport,
    ) -> Result<(), Self::Error>;

    /// Charge retained sidecar bytes/elements before copying them.
    fn charge_retained(
        &mut self,
        kind: SidecarKind,
        target: SidecarAllocation,
        amount: usize,
    ) -> Result<(), Self::Error>;

    /// Admit a sidecar list entry before retaining its value.
    fn charge_entry(
        &mut self,
        kind: SidecarKind,
        key: u32,
        source_bytes: usize,
    ) -> Result<(), Self::Error>;

    /// Return the maximum number of list entries that may be staged.
    fn max_entries(&self, kind: SidecarKind) -> usize;

    /// Map a neutral sidecar failure into the adapter error vocabulary.
    fn map_issue(&mut self, issue: SidecarIssue) -> Self::Error;

    /// Build strict comment-storage options from the current residual ledger.
    fn comment_options(
        &mut self,
        source: &[u8],
    ) -> Result<comment_storage_codec::DecodeOptions, Self::Error>;

    /// Charge one completed comment-storage codec report.
    fn charge_comment_report(
        &mut self,
        report: comment_storage_codec::DecodeReport,
    ) -> Result<(), Self::Error>;

    /// Build strict formula options from the current residual ledger.
    fn formula_options(
        &mut self,
        source: &[u8],
    ) -> Result<numbers_formula_codec::DecodeOptions, Self::Error>;

    /// Charge one completed formula codec report.
    fn charge_formula_report(
        &mut self,
        report: numbers_formula_codec::DecodeReport,
    ) -> Result<(), Self::Error>;

    /// Supply the residual aggregate limits for a strict formula-envelope
    /// admission pass.  The caller must derive these values from the same
    /// ledger used by the list and formula codec hooks.
    fn formula_envelope_limits(
        &mut self,
        source: &[u8],
    ) -> Result<formula_envelope::FormulaEnvelopeLimits, Self::Error>;

    /// Charge a successful formula-envelope preflight before its source bytes
    /// are retained.  A format adapter may use the report's wire preflight to
    /// preserve the exact aggregate field/work accounting it uses elsewhere.
    fn charge_formula_envelope_report(
        &mut self,
        report: formula_envelope::FormulaEnvelopeReport,
    ) -> Result<(), Self::Error>;

    /// Retain the wire cost spent by a rejected envelope.  Rejected source
    /// bytes are never published, but their bounded validation work remains a
    /// monotonic cost in the current read transaction.
    fn retain_formula_envelope_cost(
        &mut self,
        cost: formula_envelope::AttemptedFormulaEnvelopeCost,
    );

    /// Map a strict formula-envelope error into the adapter vocabulary.
    fn map_formula_envelope_error(&mut self, error: litchi_iwa_common::Error) -> Self::Error;
}

/// Decode one selected root/segment sidecar list with strict topology checks.
pub fn read_sidecar_list<'source, Root, Segments, Resolve, Budget>(
    kind: SidecarKind,
    object_id: u64,
    root_messages: Root,
    mut resolve_segment: Resolve,
    budget: &mut Budget,
) -> Result<SidecarList, Budget::Error>
where
    Root: IntoIterator<Item = Message<'source>>,
    Segments: IntoIterator<Item = Message<'source>>,
    Resolve: FnMut(u64, &mut Budget) -> Result<Option<Segments>, Budget::Error>,
    Budget: SidecarReadBudget,
{
    let max_entries = budget.max_entries(kind);
    let mut decoder = SidecarListDecoder { kind, budget };
    let values = table_data_list::read_list_with_decoder(
        object_id,
        kind.list_type(),
        root_messages,
        |segment_id, decoder| resolve_segment(segment_id, decoder.budget),
        &mut decoder,
        table_data_list::ListReadPolicy {
            max_entries,
            ..table_data_list::ListReadPolicy::default()
        },
    )?;
    SidecarList::from_entries(kind, values).map_err(|issue| decoder.budget.map_issue(issue))
}

struct SidecarListDecoder<'budget, Budget> {
    kind: SidecarKind,
    budget: &'budget mut Budget,
}

impl<'source, Budget> table_data_list::ListDecoder<'source> for SidecarListDecoder<'_, Budget>
where
    Budget: SidecarReadBudget,
{
    type Value = SidecarValue;
    type Error = Budget::Error;

    fn probe(
        &mut self,
        source: &'source [u8],
        kind: RootOrSegment,
        _object_id: u64,
    ) -> Result<i32, Self::Error> {
        let options = self.budget.list_options(self.kind, source)?;
        let decoded = match kind {
            RootOrSegment::Root => {
                storage::decode_table_data_list_type_with_report(source, options)
            },
            RootOrSegment::Segment => {
                storage::decode_table_data_list_segment_type_with_report(source, options)
            },
        };
        let (snapshot, report) = decoded.map_err(|error| {
            self.budget.map_issue(SidecarIssue::StorageDecode {
                kind: self.kind,
                error,
            })
        })?;
        self.budget.charge_list_probe_report(self.kind, report)?;
        Ok(snapshot.list_type())
    }

    fn decode(
        &mut self,
        source: &'source [u8],
        kind: RootOrSegment,
        _object_id: u64,
        admit: bool,
    ) -> Result<table_data_list::Candidate<Self::Value, Self::Error>, Self::Error> {
        let options = self.budget.list_options(self.kind, source)?;
        let mut visitor = SidecarListVisitor::new(self.kind, admit, self.budget);
        let decoded = match kind {
            RootOrSegment::Root => {
                storage::decode_table_data_list_with_visitor(source, options, &mut visitor)
                    .map(|(snapshot, report)| (snapshot.list_type(), None, report))
            },
            RootOrSegment::Segment => {
                storage::decode_table_data_list_segment_with_visitor(source, options, &mut visitor)
                    .map(|(snapshot, report)| {
                        (
                            snapshot.list_type(),
                            Some((snapshot.key_range_location(), snapshot.key_range_length())),
                            report,
                        )
                    })
            },
        };
        let decoded = match decoded {
            Ok(decoded) => decoded,
            Err(error) => {
                drop(visitor);
                return Err(self.budget.map_issue(SidecarIssue::StorageDecode {
                    kind: self.kind,
                    error,
                }));
            },
        };
        let (list_type, key_range, report) = decoded;
        let candidate = visitor.take_candidate(list_type, key_range);
        if admit {
            self.budget.charge_list_report(self.kind, report)?;
        } else {
            self.budget.charge_list_probe_report(self.kind, report)?;
        }
        Ok(candidate)
    }

    fn map_issue(&mut self, issue: table_data_list::CoordinatorIssue) -> Self::Error {
        self.budget.map_issue(SidecarIssue::Coordinator(issue))
    }
}

struct SidecarListVisitor<'budget, Budget: SidecarReadBudget> {
    kind: SidecarKind,
    admit: bool,
    budget: &'budget mut Budget,
    values: Vec<(u32, SidecarValue)>,
    keys: HashSet<u32>,
    segment_refs: Vec<u64>,
    segment_ids: HashSet<u64>,
    key_min: Option<u32>,
    key_max: Option<u32>,
    structural_error: Option<<Budget as SidecarReadBudget>::Error>,
    semantic_error: Option<<Budget as SidecarReadBudget>::Error>,
}

impl<'budget, Budget> SidecarListVisitor<'budget, Budget>
where
    Budget: SidecarReadBudget,
{
    fn new(kind: SidecarKind, admit: bool, budget: &'budget mut Budget) -> Self {
        Self {
            kind,
            admit,
            budget,
            values: Vec::new(),
            keys: HashSet::new(),
            segment_refs: Vec::new(),
            segment_ids: HashSet::new(),
            key_min: None,
            key_max: None,
            structural_error: None,
            semantic_error: None,
        }
    }

    fn record_structural(&mut self, issue: SidecarIssue) {
        if self.structural_error.is_none() {
            self.structural_error = Some(self.budget.map_issue(issue));
        }
    }

    fn record_semantic(&mut self, issue: SidecarIssue) {
        if self.semantic_error.is_none() {
            self.semantic_error = Some(self.budget.map_issue(issue));
        }
    }

    fn record_budget_error(&mut self, error: Budget::Error) {
        if self.semantic_error.is_none() {
            self.semantic_error = Some(error);
        }
    }

    fn take_candidate(
        self,
        list_type: i32,
        key_range: Option<(u32, u32)>,
    ) -> table_data_list::Candidate<SidecarValue, <Budget as SidecarReadBudget>::Error> {
        let entry_bounds = self
            .key_min
            .zip(self.key_max)
            .map(|(minimum, maximum)| table_data_list::EntryBounds { minimum, maximum });
        table_data_list::Candidate {
            list_type,
            values: self.values,
            keys: self.keys,
            segment_refs: self.segment_refs,
            key_range: key_range
                .map(|(location, length)| table_data_list::KeyRange { location, length }),
            entry_bounds,
            structural_error: self.structural_error,
            semantic_error: self.semantic_error,
        }
    }
}

impl<Budget> storage::StorageVisitor for SidecarListVisitor<'_, Budget>
where
    Budget: SidecarReadBudget,
{
    fn visit_list_entry_record(
        &mut self,
        record: storage::TableDataListEntryRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        let entry = record.snapshot();
        if let Some(minimum) = &mut self.key_min {
            *minimum = (*minimum).min(entry.key());
        } else {
            self.key_min = Some(entry.key());
        }
        if let Some(maximum) = &mut self.key_max {
            *maximum = (*maximum).max(entry.key());
        } else {
            self.key_max = Some(entry.key());
        }

        if !entry_matches_kind(entry, self.kind) {
            self.record_structural(SidecarIssue::InvalidEntry { kind: self.kind });
            return Ok(());
        }
        // Wrong-type and duplicate candidates still take the complete strict
        // codec path, but they never stage semantic keys or values.  This is
        // the cheap routing pass promised by the list coordinator.
        if !self.admit {
            return Ok(());
        }
        if self.kind == SidecarKind::Comments && entry.ref_count() == 0 {
            self.record_structural(SidecarIssue::InvalidEntry { kind: self.kind });
            return Ok(());
        }
        if self.keys.contains(&entry.key()) {
            self.record_structural(SidecarIssue::DuplicateKey { kind: self.kind });
            return Ok(());
        }
        let maximum = self.budget.max_entries(self.kind);
        if self.keys.len() >= maximum {
            self.record_semantic(SidecarIssue::Coordinator(
                table_data_list::CoordinatorIssue::EntryLimit {
                    observed: self.keys.len().saturating_add(1),
                    maximum,
                },
            ));
            return Ok(());
        }
        if let Err(error) = self
            .budget
            .charge_retained(self.kind, SidecarAllocation::Keys, 1)
        {
            self.record_budget_error(error);
            return Ok(());
        }
        if self.keys.try_reserve(1).is_err() {
            self.record_semantic(SidecarIssue::Allocation {
                kind: self.kind,
                target: SidecarAllocation::Keys,
                amount: self.keys.len().saturating_add(1),
            });
            return Ok(());
        }
        self.keys.insert(entry.key());
        if self.semantic_error.is_some() {
            return Ok(());
        }
        if let Err(error) = self
            .budget
            .charge_entry(self.kind, entry.key(), record.raw().len())
        {
            self.record_budget_error(error);
            return Ok(());
        }
        if let Err(error) = self
            .budget
            .charge_retained(self.kind, SidecarAllocation::Values, 1)
        {
            self.record_budget_error(error);
            return Ok(());
        }
        if self.values.try_reserve(1).is_err() {
            self.record_semantic(SidecarIssue::Allocation {
                kind: self.kind,
                target: SidecarAllocation::Values,
                amount: self.values.len().saturating_add(1),
            });
            return Ok(());
        }
        match copy_entry(self.kind, entry, self.budget) {
            Ok(value) => self.values.push((entry.key(), value)),
            Err(error) => self.record_budget_error(error),
        }
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        reference: storage::ReferenceRecord<'_>,
    ) -> Result<(), storage::DecodeError> {
        if !self.admit {
            return Ok(());
        }
        let reference = reference.reference();
        if reference.deprecated_type().is_some() || reference.deprecated_is_external().is_some() {
            self.record_structural(SidecarIssue::InvalidEntry { kind: self.kind });
            return Ok(());
        }
        let Some(identifier) = NonZeroU64::new(reference.identifier()) else {
            self.record_structural(SidecarIssue::ZeroReference { kind: self.kind });
            return Ok(());
        };
        let identifier = identifier.get();
        if self.segment_ids.contains(&identifier) {
            self.record_structural(SidecarIssue::DuplicateSegmentReference { kind: self.kind });
            return Ok(());
        }
        if let Err(error) = self
            .budget
            .charge_retained(self.kind, SidecarAllocation::Segments, 1)
        {
            self.record_budget_error(error);
            return Ok(());
        }
        if self.segment_ids.try_reserve(1).is_err() {
            self.record_semantic(SidecarIssue::Allocation {
                kind: self.kind,
                target: SidecarAllocation::Segments,
                amount: self.segment_ids.len().saturating_add(1),
            });
            return Ok(());
        }
        if self.segment_refs.try_reserve(1).is_err() {
            self.record_semantic(SidecarIssue::Allocation {
                kind: self.kind,
                target: SidecarAllocation::Segments,
                amount: self.segment_refs.len().saturating_add(1),
            });
            return Ok(());
        }
        self.segment_ids.insert(identifier);
        self.segment_refs.push(identifier);
        Ok(())
    }
}

fn entry_matches_kind(entry: storage::TableDataListEntrySnapshot<'_>, kind: SidecarKind) -> bool {
    let payloads = [
        entry.string_value().is_some(),
        entry.reference().is_some(),
        entry.formula().is_some(),
        entry.format().is_some(),
        entry.custom_format().is_some(),
        entry.rich_text_payload().is_some(),
        entry.comment_storage().is_some(),
        entry.import_warning_set().is_some(),
        entry.cell_spec().is_some(),
    ];
    if payloads.into_iter().filter(|present| *present).count() != 1 {
        return false;
    }
    match kind {
        SidecarKind::Strings | SidecarKind::FormulaErrors => entry.string_value().is_some(),
        SidecarKind::Formulas => entry.formula().is_some(),
        SidecarKind::RichTextPayloads => entry.rich_text_payload().is_some(),
        SidecarKind::Comments => entry.comment_storage().is_some(),
    }
}

fn copy_entry<Budget>(
    kind: SidecarKind,
    entry: storage::TableDataListEntrySnapshot<'_>,
    budget: &mut Budget,
) -> Result<SidecarValue, Budget::Error>
where
    Budget: SidecarReadBudget,
{
    match kind {
        SidecarKind::Strings | SidecarKind::FormulaErrors => {
            let source = entry
                .string_value()
                .ok_or_else(|| budget.map_issue(SidecarIssue::InvalidEntry { kind }))?;
            budget.charge_retained(kind, SidecarAllocation::Text, source.len())?;
            let mut value = String::new();
            value.try_reserve_exact(source.len()).map_err(|_| {
                budget.map_issue(SidecarIssue::Allocation {
                    kind,
                    target: SidecarAllocation::Text,
                    amount: source.len(),
                })
            })?;
            value.push_str(source);
            Ok(SidecarValue::Text(value.into_boxed_str()))
        },
        SidecarKind::Formulas => {
            let source = entry
                .formula()
                .ok_or_else(|| budget.map_issue(SidecarIssue::InvalidEntry { kind }))?;
            let limits = budget.formula_envelope_limits(source)?;
            match formula_envelope::preflight_formula_envelope(source, limits) {
                Ok(report) if source.is_empty() || report.root_ast_present() => {
                    budget.charge_formula_envelope_report(report)?;
                },
                Ok(report) => {
                    budget.retain_formula_envelope_cost(report.cost());
                    return Err(budget.map_formula_envelope_error(
                        litchi_iwa_common::Error::InvalidFormat(
                            "Numbers formula envelope has no root AST".to_owned(),
                        ),
                    ));
                },
                Err(failure) => {
                    let (error, cost) = failure.into_parts();
                    budget.retain_formula_envelope_cost(cost);
                    return Err(budget.map_formula_envelope_error(error));
                },
            }
            budget.charge_retained(kind, SidecarAllocation::FormulaBytes, source.len())?;
            let mut value = Vec::new();
            value.try_reserve_exact(source.len()).map_err(|_| {
                budget.map_issue(SidecarIssue::Allocation {
                    kind,
                    target: SidecarAllocation::FormulaBytes,
                    amount: source.len(),
                })
            })?;
            value.extend_from_slice(source);
            Ok(SidecarValue::Formula(value.into_boxed_slice()))
        },
        SidecarKind::RichTextPayloads | SidecarKind::Comments => {
            let reference = match kind {
                SidecarKind::RichTextPayloads => entry.rich_text_payload(),
                SidecarKind::Comments => entry.comment_storage(),
                SidecarKind::Strings | SidecarKind::Formulas | SidecarKind::FormulaErrors => None,
            };
            if let Some(reference) = reference
                && (reference.deprecated_type().is_some()
                    || reference.deprecated_is_external().is_some())
            {
                return Err(budget.map_issue(SidecarIssue::InvalidEntry { kind }));
            }
            let reference =
                reference.and_then(|reference| SidecarReference::new(reference.identifier()));
            let Some(reference) = reference else {
                return Err(budget.map_issue(SidecarIssue::ZeroReference { kind }));
            };
            Ok(SidecarValue::Reference(reference))
        },
    }
}

/// A comment-storage value with its direct reply identities retained for the
/// owning package resolver.
#[derive(Clone, Debug, PartialEq)]
pub struct CommentStorage {
    text: Box<str>,
    creation_timestamp: Option<CommentTimestamp>,
    author: Option<SidecarReference>,
    replies: Box<[SidecarReference]>,
    storage_uuid: Option<StorageUuid>,
}

impl CommentStorage {
    /// Return the comment text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return the optional Apple-epoch creation timestamp.
    #[must_use]
    pub const fn creation_date_seconds(&self) -> Option<f64> {
        match self.creation_timestamp {
            Some(timestamp) => Some(timestamp.as_f64()),
            None => None,
        }
    }

    /// Return the validated, finite Apple-epoch creation timestamp.
    #[must_use]
    pub const fn creation_timestamp(&self) -> Option<CommentTimestamp> {
        self.creation_timestamp
    }

    /// Return the optional author payload reference.
    #[must_use]
    pub const fn author_reference(&self) -> Option<SidecarReference> {
        self.author
    }

    /// Return direct reply payload references in source order.
    #[must_use]
    pub fn replies(&self) -> &[SidecarReference] {
        &self.replies
    }

    /// Return the optional storage UUID.
    #[must_use]
    pub const fn storage_uuid(&self) -> Option<StorageUuid> {
        self.storage_uuid
    }

    /// Consume the storage and expose every retained semantic and resolver
    /// part.  The author/reply references and storage UUID are deliberately
    /// returned instead of being discarded by a scalar conversion.
    #[must_use]
    pub fn into_parts(
        self,
    ) -> (
        Box<str>,
        Option<CommentTimestamp>,
        Option<SidecarReference>,
        Box<[SidecarReference]>,
        Option<StorageUuid>,
    ) {
        (
            self.text,
            self.creation_timestamp,
            self.author,
            self.replies,
            self.storage_uuid,
        )
    }

    /// Convert the scalar comment fields to the neutral table comment model.
    ///
    /// The package-owned author/reply references are available through
    /// [`Self::into_parts`] for adapters that resolve them.  This scalar
    /// conversion intentionally produces a comment without those unresolved
    /// native references.
    pub fn try_into_comment(self) -> Result<Comment, litchi_iwa_common::table::model::Error> {
        // `Box<str>` converts into `String` by moving its allocation; using
        // the infallible constructor here avoids copying the already bounded
        // comment text a second time.
        Ok(Comment::with_metadata(
            String::from(self.text),
            self.creation_timestamp,
            None,
        ))
    }
}

/// A source UUID retained by comment storage.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct StorageUuid {
    lower: u64,
    upper: u64,
}

impl StorageUuid {
    /// Construct a UUID, preserving all source bits.
    #[must_use]
    pub const fn new(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }

    /// Return the lower 64 bits.
    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }

    /// Return the upper 64 bits.
    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

struct ReplyVisitor<'budget, Budget: SidecarReadBudget> {
    root: SidecarReference,
    kind: SidecarKind,
    budget: &'budget mut Budget,
    replies: Vec<SidecarReference>,
    seen: HashSet<SidecarReference>,
    issue: bool,
    budget_error: Option<<Budget as SidecarReadBudget>::Error>,
}

impl<'budget, Budget> ReplyVisitor<'budget, Budget>
where
    Budget: SidecarReadBudget,
{
    fn new(root: SidecarReference, budget: &'budget mut Budget) -> Self {
        Self {
            root,
            kind: SidecarKind::Comments,
            budget,
            replies: Vec::new(),
            seen: HashSet::new(),
            issue: false,
            budget_error: None,
        }
    }
}

impl<Budget> comment_storage_codec::CommentStorageVisitor for ReplyVisitor<'_, Budget>
where
    Budget: SidecarReadBudget,
{
    fn visit_reply(
        &mut self,
        reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), comment_storage_codec::DecodeError> {
        if self.budget_error.is_some() {
            return Ok(());
        }
        let reply = reply.reference();
        if reply.deprecated_type().is_some() || reply.deprecated_is_external().is_some() {
            self.issue = true;
            return Ok(());
        }
        let Some(reference) = SidecarReference::new(reply.identifier()) else {
            self.issue = true;
            return Ok(());
        };
        if reference == self.root || self.seen.contains(&reference) {
            self.issue = true;
            return Ok(());
        }
        if let Err(error) = self
            .budget
            .charge_retained(self.kind, SidecarAllocation::Comments, 1)
        {
            self.budget_error = Some(error);
            return Ok(());
        }
        if self.replies.try_reserve(1).is_err() {
            self.budget_error = Some(self.budget.map_issue(SidecarIssue::Allocation {
                kind: self.kind,
                target: SidecarAllocation::Comments,
                amount: self.replies.len().saturating_add(1),
            }));
            return Ok(());
        }
        if self.seen.try_reserve(1).is_err() {
            self.budget_error = Some(self.budget.map_issue(SidecarIssue::Allocation {
                kind: self.kind,
                target: SidecarAllocation::Comments,
                amount: self.seen.len().saturating_add(1),
            }));
            return Ok(());
        }
        // `contains` above intentionally precedes both reservations.  Insert
        // only after each fallible growth operation succeeds so a refused
        // allocation cannot leave a logical identity in the staged set.
        self.seen.insert(reference);
        self.replies.push(reference);
        Ok(())
    }
}

/// Decode one canonical comment-storage payload from an object's messages.
pub fn read_comment_storage<'source, Messages, Budget>(
    root: SidecarReference,
    messages: Messages,
    budget: &mut Budget,
) -> Result<CommentStorage, Budget::Error>
where
    Messages: IntoIterator<Item = Message<'source>>,
    Budget: SidecarReadBudget,
{
    let mut payload_count = 0usize;
    let mut duplicate = false;
    let mut first_snapshot = None;
    let mut first_replies = None;
    for message in messages {
        if message.kind != COMMENT_STORAGE_MESSAGE_KIND {
            continue;
        }
        payload_count = payload_count.saturating_add(1);
        let source = message.data;
        let options = budget.comment_options(source)?;
        if first_snapshot.is_none() {
            let mut visitor = ReplyVisitor::new(root, budget);
            let decoded = comment_storage_codec::decode_comment_storage_archive_with_visitor(
                source,
                options,
                &mut visitor,
            );
            let (snapshot, report) = match decoded {
                Ok(decoded) => decoded,
                Err(error) => {
                    drop(visitor);
                    return Err(budget.map_issue(SidecarIssue::CommentDecode(error)));
                },
            };
            let issue = visitor.issue;
            let budget_error = visitor.budget_error.take();
            let replies = std::mem::take(&mut visitor.replies);
            drop(visitor);
            budget.charge_comment_report(report)?;
            if let Some(error) = budget_error {
                return Err(error);
            }
            if issue || replies.len() != report.reply_references() {
                return Err(budget.map_issue(SidecarIssue::InvalidReplies));
            }
            first_snapshot = Some(snapshot);
            first_replies = Some(replies);
        } else {
            duplicate = true;
            // Duplicate canonical payloads still take the strict decode path
            // so malformed bytes cannot be hidden behind the duplicate issue.
            let (_, report) = comment_storage_codec::decode_comment_storage_archive_with_visitor(
                source,
                options,
                &mut (),
            )
            .map_err(|error| budget.map_issue(SidecarIssue::CommentDecode(error)))?;
            budget.charge_comment_report(report)?;
        }
    }
    if payload_count == 0 {
        return Err(budget.map_issue(SidecarIssue::MissingPayload {
            kind: SidecarKind::Comments,
        }));
    }
    if duplicate {
        return Err(budget.map_issue(SidecarIssue::DuplicatePayload {
            kind: SidecarKind::Comments,
        }));
    }
    let Some(snapshot) = first_snapshot else {
        return Err(budget.map_issue(SidecarIssue::MissingPayload {
            kind: SidecarKind::Comments,
        }));
    };
    let Some(replies) = first_replies else {
        return Err(budget.map_issue(SidecarIssue::InvalidReplies));
    };
    let text = snapshot.text().unwrap_or_default();
    budget.charge_retained(SidecarKind::Comments, SidecarAllocation::Text, text.len())?;
    let mut retained_text = String::new();
    retained_text.try_reserve_exact(text.len()).map_err(|_| {
        budget.map_issue(SidecarIssue::Allocation {
            kind: SidecarKind::Comments,
            target: SidecarAllocation::Text,
            amount: text.len(),
        })
    })?;
    retained_text.push_str(text);

    let author = match snapshot.author() {
        Some(reference)
            if reference.deprecated_type().is_some()
                || reference.deprecated_is_external().is_some() =>
        {
            return Err(budget.map_issue(SidecarIssue::InvalidEntry {
                kind: SidecarKind::Comments,
            }));
        },
        Some(reference) => Some(SidecarReference::new(reference.identifier()).ok_or_else(
            || {
                budget.map_issue(SidecarIssue::ZeroReference {
                    kind: SidecarKind::Comments,
                })
            },
        )?),
        None => None,
    };
    let storage_uuid = match snapshot.storage_uuid() {
        Some(uuid) if uuid.lower() == 0 && uuid.upper() == 0 => {
            return Err(budget.map_issue(SidecarIssue::ZeroUuid));
        },
        Some(uuid) => Some(StorageUuid::new(uuid.lower(), uuid.upper())),
        None => None,
    };
    let creation_timestamp = match snapshot.creation_date() {
        Some(date) => Some(CommentTimestamp::new(date.seconds()).ok_or_else(|| {
            budget.map_issue(SidecarIssue::InvalidEntry {
                kind: SidecarKind::Comments,
            })
        })?),
        None => None,
    };
    Ok(CommentStorage {
        text: retained_text.into_boxed_str(),
        creation_timestamp,
        author,
        replies: replies.into_boxed_slice(),
        storage_uuid,
    })
}

/// Formula codec budget hooks layered on top of the sidecar ledger.
pub trait FormulaSidecarBudget:
    SidecarReadBudget + FormulaEventRenderBudget<Error = <Self as SidecarReadBudget>::Error>
{
}

impl<T> FormulaSidecarBudget for T where
    T: SidecarReadBudget + FormulaEventRenderBudget<Error = <T as SidecarReadBudget>::Error>
{
}

/// Render one retained formula archive through the shared compatibility
/// renderer.  The archive is decoded only when a formula cell is visited.
pub fn render_formula<R, Budget>(
    source: &[u8],
    owner: u32,
    row: u32,
    column: u32,
    rows: u32,
    columns: u32,
    resolver: &R,
    budget: &mut Budget,
) -> Result<String, <Budget as SidecarReadBudget>::Error>
where
    R: ReferenceResolver,
    Budget: FormulaSidecarBudget,
{
    let options = budget.formula_options(source)?;
    let context = numbers_formula_codec::FormulaContext::new(owner, row, column, rows, columns);
    let mut visitor = formula_render::FormulaRenderCodecVisitor::new(row, column, resolver, budget);
    let decoded = numbers_formula_codec::decode_formula_archive_for_render(
        source,
        context,
        options,
        &mut visitor,
    );
    match decoded {
        Ok(report) => {
            visitor.budget_mut().charge_formula_report(report)?;
            if let Some(error) = visitor.take_error() {
                return Err(error);
            }
            visitor.finish()
        },
        Err(error) => {
            if let Some(error) = visitor.take_error() {
                return Err(error);
            }
            Err(budget.map_issue(SidecarIssue::FormulaDecode(error)))
        },
    }
}

/// Hooks for resolving payloads while converting one cell classification into
/// the common semantic value model.
pub trait CellSidecarResolver {
    /// Adapter error type.
    type Error;

    /// Resolve one rich-text object to owned display text.
    fn rich_text(&mut self, reference: SidecarReference) -> Result<String, Self::Error>;

    /// Render one formula source archive.
    fn formula(
        &mut self,
        key: u32,
        source: &[u8],
        row: u32,
        column: u32,
    ) -> Result<String, Self::Error>;

    /// Resolve one comment-storage object into the common comment value.
    fn comment(&mut self, reference: SidecarReference) -> Result<Comment, Self::Error>;

    /// Retain a source text value while charging the adapter's cumulative
    /// output ledger.  The helper never calls `to_owned` on semantic text
    /// without passing through this hook.
    fn retain_text(
        &mut self,
        kind: SidecarKind,
        key: u32,
        source: &str,
    ) -> Result<String, Self::Error>;

    /// Construct an error for a missing or malformed sidecar entry.
    fn missing(&mut self, kind: SidecarKind, key: u32) -> Self::Error;

    /// Construct an error for a non-finite semantic scalar.
    fn invalid_scalar(&mut self) -> Self::Error;
}

impl SidecarTables {
    /// Resolve one source-classified cell into a common value and optional
    /// comment using the adapter hooks.
    pub fn materialize_cell<Resolver>(
        &self,
        source: CellValueSource,
        row: u32,
        column: u32,
        resolver: &mut Resolver,
    ) -> Result<(Value, Option<Comment>), Resolver::Error>
    where
        Resolver: CellSidecarResolver,
    {
        let value = match source.value {
            ValueSource::Empty => Value::Empty,
            ValueSource::Number(value) => {
                Value::Number(FiniteF64::new(value.get()).map_err(|_| resolver.invalid_scalar())?)
            },
            ValueSource::Date(value) => {
                Value::Date(FiniteF64::new(value.get()).map_err(|_| resolver.invalid_scalar())?)
            },
            ValueSource::Boolean(value) => Value::Boolean(value),
            ValueSource::Duration(value) => {
                Value::Duration(FiniteF64::new(value.get()).map_err(|_| resolver.invalid_scalar())?)
            },
            ValueSource::Text(key) => {
                let Some(text) = self.string(key) else {
                    // A missing string sidecar entry has historically
                    // represented an empty stored cell.  Preserve that
                    // compatibility distinction; formula/comment references
                    // remain strict below.
                    return finish_cell(self, resolver, Value::Empty, source.comment_identifier);
                };
                Value::Text(resolver.retain_text(SidecarKind::Strings, key, text)?)
            },
            ValueSource::RichText(key) => {
                let Some(reference) = self.rich_text_reference(key) else {
                    return finish_cell(self, resolver, Value::Empty, source.comment_identifier);
                };
                Value::Text(resolver.rich_text(reference)?)
            },
            ValueSource::Formula(key) => {
                let formula = self
                    .formula(key)
                    .ok_or_else(|| resolver.missing(SidecarKind::Formulas, key))?;
                Value::Formula(resolver.formula(key, formula, row, column)?)
            },
            ValueSource::Error(key) => {
                let text = key
                    .and_then(|key| self.formula_error(key))
                    .unwrap_or("FORMULA");
                Value::Error(resolver.retain_text(
                    SidecarKind::FormulaErrors,
                    key.unwrap_or_default(),
                    text,
                )?)
            },
        };
        finish_cell(self, resolver, value, source.comment_identifier)
    }
}

fn finish_cell<Resolver>(
    tables: &SidecarTables,
    resolver: &mut Resolver,
    value: Value,
    comment_key: Option<u32>,
) -> Result<(Value, Option<Comment>), Resolver::Error>
where
    Resolver: CellSidecarResolver,
{
    let comment = comment_key
        .map(|key| {
            let reference = tables
                .comment_reference(key)
                .ok_or_else(|| resolver.missing(SidecarKind::Comments, key))?;
            resolver.comment(reference)
        })
        .transpose()?;
    Ok((value, comment))
}

/// Construct one common cell-comment record for a selected position.
pub fn positioned_comment(
    position: litchi_iwa_common::table::coordinate::CellPosition,
    comment: Comment,
) -> CellComment {
    CellComment::new(position, comment)
}
