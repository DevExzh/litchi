//! Guarded source-backed merged-range edits for one existing worksheet.
//!
//! This owner keeps the merge operation deliberately narrow.  It captures the
//! workbook catalog and one selected worksheet, validates merge semantics with
//! the ordinary worksheet parser and merge rewriter, and publishes one
//! worksheet-Part overlay.  The OPC package and XML representation remain
//! implementation details of this module.

use std::io::Write;
use std::sync::Arc;

use litchi_core::{
    ExecutionContext, ExecutionError, ReadAt, Selector as CoreSelector, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    OpcError, OpcPackage, PackURI, Part, PartView, ReadLimits, Relationship, Relationships,
    SourceBackedPackage, SourceCacheLimits, SourceLineage, TargetMode,
};
use litchi_sheet::{Area, At, Cell as Address, Rect};
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};

use super::source::validate_sheet_graph;
use crate::cell::{Cell, Store, Text};
use crate::error::{Error, MergeEditBlock, Result, allocation, invalid};
use crate::merge;
use crate::raw;
use crate::source_payload::SourcePayload;
use crate::{Selector, WorksheetKind};

/// An owning source-backed merged-range editor for one XLSX artifact.
///
/// The editor retains the deferred OPC source but exposes only semantic
/// worksheet operations.  A changed commit replaces exactly one existing
/// worksheet payload; every other physical member is copied by the OPC owner.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
}

/// An isolated merged-range edit over one exact source worksheet.
pub struct SourceEdit {
    before: Snapshot,
    staged: Vec<Rect>,
    store: Store,
    shared_strings: Option<Arc<[Text]>>,
}

/// Immutable merged-range state and the private source closure required for
/// exact publication.
#[derive(Clone, Debug)]
pub struct Snapshot {
    ranges: Box<[Rect]>,
    sheet_name: Box<str>,
    sheet_position: usize,
    source: SourceState,
}

/// Exact reversible source-bound merged-range replacement.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

/// Successful source-backed merged-range publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    diagnostics: Diagnostics,
}

/// Content-free source-backed merge publication diagnostics.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    touched_worksheets: u8,
}

impl SourceBackedEditor {
    /// Open with the standard bounded OPC policy.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open with an explicit bounded OPC policy.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source,
            read_limits,
        )?)
    }

    /// Open with an explicit finite deferred-payload cache policy.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_cache_limits(
            source,
            cache_limits,
        )?)
    }

    /// Open with explicit read and cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                read_limits,
                cache_limits,
            )?,
        )
    }

    /// Open with an explicit managed execution context.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_execution_context(
            source,
            read_limits,
            context,
        )?)
    }

    /// Open with explicit read and managed execution policies.
    pub fn from_read_at_with_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(source, read_limits, context)
    }

    /// Open with explicit read, cache, and managed execution policies.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                read_limits,
                cache_limits,
                context,
            )?,
        )
    }

    /// Build an editor from an already opened deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        package.check_execution()?;
        Ok(Self { package })
    }

    /// Capture exact source-bound merge state for one selected worksheet.
    pub fn snapshot<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<Snapshot>> {
        self.package.check_execution()?;
        Ok(Snapshot::load_source_backed(&self.package, selector)?.map(|loaded| loaded.0))
    }

    /// Begin an isolated source-backed edit without loading unselected
    /// worksheet payloads.
    pub fn edit<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<SourceEdit>> {
        self.package.check_execution()?;
        let Some((snapshot, store, shared_strings)) =
            Snapshot::load_source_backed(&self.package, selector)?
        else {
            return Ok(None);
        };
        Ok(Some(SourceEdit::new(snapshot, store, shared_strings)))
    }

    /// Publish a source-checked commit to a sequential sink.
    ///
    /// A changed commit emits one replacement for the selected worksheet Part.
    /// Exact no-ops copy the complete source artifact byte-for-byte.
    pub fn publish_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> Result<Snapshot> {
        self.package.check_execution()?;
        if !commit.patch.before.matches_source_backed(&self.package)? {
            return Err(Error::PatchConflict {
                part: commit.patch.before.part_name().to_string(),
            });
        }
        let target = if commit.patch.is_empty() {
            commit.patch.before.clone()
        } else {
            commit.patch.after.clone()
        };
        target.check_execution()?;
        if commit.patch.is_empty() {
            self.package
                .write_part_overlays_shared_to_stream(writer, Vec::new())?;
        } else {
            self.package.write_part_overlay_shared_to_stream(
                writer,
                target.part_name(),
                target.source_arc()?,
            )?;
        }
        Ok(target)
    }
}

impl SourceEdit {
    fn new(before: Snapshot, store: Store, shared_strings: Option<Arc<[Text]>>) -> Self {
        Self {
            staged: before.ranges.to_vec(),
            before,
            store,
            shared_strings,
        }
    }

    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged merged ranges in deterministic worksheet order.
    #[must_use]
    pub fn merges(&self) -> merge::Merges<'_> {
        merge::Merges::new(&self.staged)
    }

    /// Stage a checked rectangular merge without discarding follower cells.
    pub fn merge<'a>(&mut self, area: impl Into<Area<'a>>) -> Result<&mut Self> {
        self.before.check_execution()?;
        let range = area.into().resolve()?;
        ensure_merge_area(self.before.sheet_name(), range)?;
        if self.staged.contains(&range) {
            return Ok(self);
        }
        self.staged
            .try_reserve(1)
            .map_err(|source| allocation("source-backed merge edit plan", source))?;
        self.staged.push(range);
        self.staged.sort_unstable_by_key(canonical_range_key);
        Ok(self)
    }

    /// Stage removal of the merged range containing one checked cell.
    pub fn unmerge<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        self.before.check_execution()?;
        let address = at.into().resolve()?;
        let Some(range) = self
            .staged
            .iter()
            .copied()
            .find(|range| range.contains(address))
        else {
            return Ok(self);
        };
        self.staged.retain(|candidate| *candidate != range);
        Ok(self)
    }

    /// Whether the authored merged-range state differs from its source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.ranges.as_ref() != self.staged.as_slice()
    }

    /// Validate and freeze this isolated edit for source-backed publication.
    pub fn commit(self) -> Result<Commit> {
        self.before.check_execution()?;
        let (add, remove) = plan_changes(&self.before.ranges, &self.staged)?;
        if add.is_empty() && remove.is_empty() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }

        validate_structural_preconditions(&self.before, &add, &remove)?;
        let output = raw::worksheet::edit::rewrite_merges(
            self.before.source_xml(),
            self.before.sheet_name(),
            raw::worksheet::edit::MergePlan { add, remove },
        )?;
        validate_follower_content(&self.before, &self.store, &self.staged)?;
        let ranges = canonical_ranges(self.staged)?;
        let snapshot = Snapshot::from_rewritten_source(
            &self.before,
            output,
            ranges,
            self.shared_strings.as_deref(),
        )?;
        snapshot.check_execution()?;
        let patch = Patch::new(self.before, snapshot.clone());
        Ok(Commit::new(snapshot, patch, true))
    }
}

impl Snapshot {
    /// Number of direct merged ranges in the selected worksheet.
    #[must_use]
    pub fn merge_count(&self) -> usize {
        self.ranges.len()
    }

    /// Lazily iterate direct merged ranges in deterministic order.
    #[must_use]
    pub fn merges(&self) -> merge::Merges<'_> {
        merge::Merges::new(&self.ranges)
    }

    /// Find the direct merged range containing one checked address.
    #[must_use]
    pub fn merge_at(&self, address: Address) -> Option<Rect> {
        self.ranges
            .iter()
            .copied()
            .find(|range| range.contains(address))
    }

    /// Whether the selected worksheet contains the exact direct merged range.
    #[must_use]
    pub fn contains_merge(&self, range: Rect) -> bool {
        self.ranges.contains(&range)
    }

    /// Developer-facing worksheet name captured at ingress.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Checked zero-based worksheet position captured at ingress.
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.sheet_position
    }

    fn load(package: &OpcPackage, sheet_position: usize) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let workbook_xml = workbook.blob();
        let catalog = raw::parse_catalog(workbook_xml)?;
        let catalog_sheet = catalog
            .sheets
            .get(sheet_position)
            .ok_or_else(|| invalid("merged-range worksheet position is absent"))?;
        let worksheet_uri = workbook
            .rels()
            .get(&catalog_sheet.relationship_id)
            .ok_or_else(|| invalid("selected worksheet relationship is missing"))?
            .target_partname()?;
        let worksheet = package.get_part(&worksheet_uri)?;
        let sheet_relationship = require_selected_worksheet(
            workbook.rels(),
            catalog_sheet,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let owner_relationship = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;
        let (store, _shared_strings, shared_strings_source) =
            parse_owned_worksheet_store(package, workbook, worksheet.blob())?;
        let ranges = canonical_ranges(store.merge_ranges().to_vec())?;
        let source = SourceState::new(
            workbook.partname().clone(),
            workbook.content_type(),
            SourcePayload::Owned(workbook.blob_arc()),
            owner_relationship,
            worksheet.partname().clone(),
            worksheet.content_type(),
            SourcePayload::Owned(worksheet.blob_arc()),
            sheet_relationship,
            worksheet.rels(),
            workbook.rels(),
            shared_strings_source,
            None,
            None,
            None,
        )?;
        Ok(Self {
            ranges: ranges.into_boxed_slice(),
            sheet_name: copy_boxed(&catalog_sheet.name, "merged-range sheet name")?,
            sheet_position,
            source,
        })
    }

    fn load_source_backed<'a>(
        package: &SourceBackedPackage,
        selector: impl Into<Selector<'a>>,
    ) -> Result<Option<(Self, Store, Option<Arc<[Text]>>)>> {
        package.check_execution()?;
        let source_version = package.source_version()?;
        let source_lineage = package.source_lineage();
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let workbook_data = workbook.data()?;
        let workbook_xml = SourcePayload::from_part_data(package, workbook_data)?;
        let catalog = raw::parse_catalog(workbook_xml.as_bytes())?;
        let sheet_parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
        let Some(sheet_position) = resolve_selector(&catalog.sheets, selector.into())? else {
            package.check_execution()?;
            if package.source_version()? != source_version {
                return Err(invalid(
                    "merged-range source version changed during snapshot",
                ));
            }
            return Ok(None);
        };
        let catalog_sheet = catalog
            .sheets
            .get(sheet_position)
            .ok_or_else(|| invalid("merged-range worksheet position disappeared"))?;
        let sheet_part = sheet_parts
            .get(sheet_position)
            .ok_or_else(|| invalid("merged-range worksheet part disappeared"))?;
        if sheet_part.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: catalog_sheet.name.clone(),
            });
        }
        let worksheet = package.part(&sheet_part.uri)?;
        let sheet_relationship = require_selected_worksheet(
            workbook.rels(),
            catalog_sheet,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let worksheet_data = worksheet.data()?;
        let worksheet_xml = SourcePayload::from_part_data(package, worksheet_data)?;
        let (store, shared_strings, shared_strings_source) =
            parse_worksheet_store(package, &workbook, worksheet_xml.as_bytes())?;
        let mut ranges = Vec::new();
        ranges
            .try_reserve_exact(store.merge_ranges().len())
            .map_err(|source| allocation("source-backed merged ranges", source))?;
        ranges.extend_from_slice(store.merge_ranges());
        let ranges = canonical_ranges(ranges)?;
        let owner_relationship = current_owner_relationship(package.rels())
            .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?;
        let source = SourceState::new(
            workbook.partname().clone(),
            workbook.content_type(),
            workbook_xml,
            owner_relationship,
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet_xml,
            sheet_relationship,
            worksheet.rels(),
            workbook.rels(),
            shared_strings_source,
            Some(source_lineage),
            Some(source_version),
            package.execution_context(),
        )?;
        package.check_execution()?;
        if package.source_version()? != source_version {
            return Err(invalid(
                "merged-range source version changed during snapshot",
            ));
        }
        Ok(Some((
            Self {
                ranges: ranges.into_boxed_slice(),
                sheet_name: copy_boxed(&catalog_sheet.name, "source-backed merge sheet name")?,
                sheet_position,
                source,
            },
            store,
            shared_strings,
        )))
    }

    fn from_rewritten_source(
        source: &Self,
        bytes: Vec<u8>,
        ranges: Vec<Rect>,
        shared_strings: Option<&[Text]>,
    ) -> Result<Self> {
        source.check_execution()?;
        let store = parse_worksheet_store_bytes(&bytes, shared_strings)?;
        let readback = canonical_ranges(store.merge_ranges().to_vec())?;
        if readback != ranges {
            return Err(invalid(
                "merged-range rewrite readback did not match its target",
            ));
        }
        let mut rewritten = source.clone();
        rewritten.ranges = ranges.into_boxed_slice();
        rewritten.source.worksheet.bytes = SourcePayload::Owned(Arc::new(bytes));
        rewritten.check_execution()?;
        Ok(rewritten)
    }

    fn source_xml(&self) -> &[u8] {
        self.source.worksheet.bytes.as_bytes()
    }

    fn source_arc(&self) -> Result<Arc<Vec<u8>>> {
        self.source.worksheet.bytes.detached_arc()
    }

    fn materialized_source_arc(&self, maximum_bytes: usize) -> Result<Arc<Vec<u8>>> {
        self.source
            .worksheet
            .bytes
            .materialized_arc(maximum_bytes, "source-backed merge worksheet")
    }

    fn part_name(&self) -> &PackURI {
        &self.source.worksheet.uri
    }

    fn check_execution(&self) -> Result<()> {
        self.source.check_execution()
    }

    fn matches_source_backed(&self, package: &SourceBackedPackage) -> Result<bool> {
        self.source
            .matches_source_backed(package, &self.sheet_name, self.sheet_position)
    }

    fn matches_current_source(&self, package: &OpcPackage) -> bool {
        self.source
            .matches_current_source(package, &self.sheet_name, self.sheet_position)
    }
}

impl Patch {
    const fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Required exact source state.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact target state produced by the commit.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the patch preserves the selected source worksheet byte-for-byte.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact source-bound inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this exact source-bound worksheet replacement atomically.
    ///
    /// A managed source target returns a typed
    /// `ManagedPartDataArcEscape` error rather than detaching its reservation.
    /// Use [`Self::apply_materialized`] for an explicitly bounded handoff.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply_inner(package, None)
    }

    /// Apply after explicitly materializing the target worksheet under a byte bound.
    pub fn apply_materialized(&self, package: &mut OpcPackage, maximum_bytes: usize) -> Result<()> {
        self.apply_inner(package, Some(maximum_bytes))
    }

    fn apply_inner(&self, package: &mut OpcPackage, maximum_bytes: Option<usize>) -> Result<()> {
        self.before.check_execution()?;
        self.after.check_execution()?;
        if !self.before.matches_current_source(package) {
            return Err(Error::PatchConflict {
                part: self.before.part_name().to_string(),
            });
        }
        if self.is_empty() {
            self.before.check_execution()?;
            self.after.check_execution()?;
            return Ok(());
        }
        if package.is_signed() {
            return Err(Error::Signed);
        }
        let blob = match maximum_bytes {
            Some(maximum) => self.after.materialized_source_arc(maximum)?,
            None => self.after.source_arc()?,
        };
        self.after.check_execution()?;
        let mut candidate = package.clone();
        candidate
            .get_part_mut(self.before.part_name())?
            .set_blob_shared(blob);
        let resulting = Snapshot::load(&candidate, self.after.sheet_position)?;
        self.after.check_execution()?;
        if !resulting.same_source(&self.after)
            || resulting.ranges.as_ref() != self.after.ranges.as_ref()
        {
            return Err(invalid(
                "merged-range patch readback did not match its target",
            ));
        }
        *package = candidate;
        Ok(())
    }
}

impl Commit {
    const fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            diagnostics: Diagnostics {
                touched_worksheets: if changed { 1 } else { 0 },
            },
        }
    }

    /// Whether the authored merged-range state changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.diagnostics.touched_worksheets != 0
    }

    /// Resulting immutable source-bound state.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume this result into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

impl Diagnostics {
    /// Number of worksheet Parts replaced by this commit.
    #[must_use]
    pub fn touched_worksheets(self) -> usize {
        usize::from(self.touched_worksheets)
    }
}

#[derive(Clone, Debug)]
struct SourceState {
    workbook: PartState,
    worksheet: PartState,
    owner_relationship: SourceRelationship,
    workbook_relationships: Vec<SourceRelationship>,
    sheet_relationship: SourceRelationship,
    worksheet_relationships: Vec<SourceRelationship>,
    shared_strings: Option<SharedStringsState>,
    lineage: Option<SourceLineage>,
    version: Option<SourceVersion>,
    context: Option<ExecutionContext>,
}

impl SourceState {
    #[allow(clippy::too_many_arguments)]
    fn new(
        workbook_uri: PackURI,
        workbook_content_type: &str,
        workbook_bytes: SourcePayload,
        owner_relationship: &Relationship,
        worksheet_uri: PackURI,
        worksheet_content_type: &str,
        worksheet_bytes: SourcePayload,
        sheet_relationship: &Relationship,
        worksheet_relationships: &Relationships,
        workbook_relationships: &Relationships,
        shared_strings: Option<SharedStringsState>,
        lineage: Option<SourceLineage>,
        version: Option<SourceVersion>,
        context: Option<ExecutionContext>,
    ) -> Result<Self> {
        Ok(Self {
            workbook: PartState::new(
                workbook_uri,
                workbook_content_type,
                workbook_bytes,
                "source-backed merge workbook content type",
            )?,
            worksheet: PartState::new(
                worksheet_uri,
                worksheet_content_type,
                worksheet_bytes,
                "source-backed merge worksheet content type",
            )?,
            owner_relationship: SourceRelationship::capture(
                owner_relationship,
                "source-backed merge owner relationship",
            )?,
            workbook_relationships: capture_relationships(workbook_relationships)?,
            sheet_relationship: SourceRelationship::capture(
                sheet_relationship,
                "source-backed merge worksheet relationship",
            )?,
            worksheet_relationships: capture_relationships(worksheet_relationships)?,
            shared_strings,
            lineage,
            version,
            context,
        })
    }

    fn check_execution(&self) -> Result<()> {
        check_execution(self.context.as_ref())
    }

    fn same_source(&self, other: &Self) -> bool {
        self.workbook.same_source(&other.workbook)
            && self.worksheet.same_source(&other.worksheet)
            && self.owner_relationship == other.owner_relationship
            && self.workbook_relationships == other.workbook_relationships
            && self.sheet_relationship == other.sheet_relationship
            && self.worksheet_relationships == other.worksheet_relationships
            && optional_shared_strings_same(&self.shared_strings, &other.shared_strings)
            && optional_same(&self.lineage, &other.lineage)
            && optional_same(&self.version, &other.version)
    }

    fn matches_source_backed(
        &self,
        package: &SourceBackedPackage,
        sheet_name: &str,
        sheet_position: usize,
    ) -> Result<bool> {
        package.check_execution()?;
        let (Some(lineage), Some(version)) = (&self.lineage, self.version) else {
            return Ok(false);
        };
        if package.source_lineage() != lineage.clone() || package.source_version()? != version {
            return Ok(false);
        }
        let workbook = package.main_document_part()?;
        if !self
            .workbook
            .matches_view(&workbook, workbook.data()?.as_bytes())
            || !current_owner_relationship(package.rels())
                .is_some_and(|value| self.owner_relationship.matches(value))
        {
            return Ok(false);
        }
        if !relationships_match(&self.workbook_relationships, workbook.rels()) {
            return Ok(false);
        }
        if let Some(shared_strings) = self.shared_strings.as_ref()
            && !shared_strings.matches_source_backed(package, &workbook)?
        {
            return Ok(false);
        }
        let workbook_xml = workbook.data()?;
        let catalog = raw::parse_catalog(workbook_xml.as_bytes())?;
        let sheet_parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
        let Some(catalog_sheet) = catalog.sheets.get(sheet_position) else {
            return Ok(false);
        };
        let Some(sheet_part) = sheet_parts.get(sheet_position) else {
            return Ok(false);
        };
        if catalog_sheet.name != sheet_name
            || sheet_part.kind != WorksheetKind::Worksheet
            || sheet_part.uri != self.worksheet.uri
        {
            return Ok(false);
        }
        let worksheet = package.part(&sheet_part.uri)?;
        let relationship = require_selected_worksheet(
            workbook.rels(),
            catalog_sheet,
            worksheet.partname(),
            worksheet.content_type(),
        )?;
        let worksheet_xml = worksheet.data()?;
        if !self.sheet_relationship.matches(relationship)
            || !self
                .worksheet
                .matches_view(&worksheet, worksheet_xml.as_bytes())
            || !relationships_match(&self.worksheet_relationships, worksheet.rels())
        {
            return Ok(false);
        }
        package.check_execution()?;
        Ok(package.source_version()? == version)
    }

    fn matches_current_source(
        &self,
        package: &OpcPackage,
        sheet_name: &str,
        sheet_position: usize,
    ) -> bool {
        let Ok(workbook) = package.main_document_part() else {
            return false;
        };
        if !self.workbook.matches_part(workbook)
            || !current_owner_relationship(package.rels())
                .is_some_and(|value| self.owner_relationship.matches(value))
        {
            return false;
        }
        if !relationships_match(&self.workbook_relationships, workbook.rels()) {
            return false;
        }
        if let Some(shared_strings) = self.shared_strings.as_ref()
            && !shared_strings.matches_current_source(package, workbook)
        {
            return false;
        }
        let Ok(catalog) = raw::parse_catalog(workbook.blob()) else {
            return false;
        };
        let Some(catalog_sheet) = catalog.sheets.get(sheet_position) else {
            return false;
        };
        if catalog_sheet.name != sheet_name {
            return false;
        }
        let Some(relationship) = workbook.rels().get(&catalog_sheet.relationship_id) else {
            return false;
        };
        let Ok(uri) = relationship.target_partname() else {
            return false;
        };
        let Ok(worksheet) = package.get_part(&uri) else {
            return false;
        };
        let Ok(selected) = require_selected_worksheet(
            workbook.rels(),
            catalog_sheet,
            worksheet.partname(),
            worksheet.content_type(),
        ) else {
            return false;
        };
        self.sheet_relationship.matches(selected)
            && self.worksheet.matches_part(worksheet)
            && relationships_match(&self.worksheet_relationships, worksheet.rels())
    }
}

impl Snapshot {
    fn same_source(&self, other: &Self) -> bool {
        self.sheet_name == other.sheet_name
            && self.sheet_position == other.sheet_position
            && self.source.same_source(&other.source)
    }
}

#[derive(Clone, Debug)]
struct PartState {
    uri: PackURI,
    content_type: Box<str>,
    bytes: SourcePayload,
}

impl PartState {
    fn new(
        uri: PackURI,
        content_type: &str,
        bytes: SourcePayload,
        resource: &'static str,
    ) -> Result<Self> {
        Ok(Self {
            uri,
            content_type: copy_boxed(content_type, resource)?,
            bytes,
        })
    }

    fn same_source(&self, other: &Self) -> bool {
        self.uri == other.uri
            && self.content_type == other.content_type
            && self.bytes == other.bytes
    }

    fn matches_view(&self, part: &PartView<'_>, bytes: &[u8]) -> bool {
        part.partname() == &self.uri
            && part.content_type() == self.content_type.as_ref()
            && bytes == self.bytes.as_bytes()
    }

    fn matches_part(&self, part: &dyn Part) -> bool {
        part.partname() == &self.uri
            && part.content_type() == self.content_type.as_ref()
            && part.blob() == self.bytes.as_bytes()
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceRelationship {
    id: Box<str>,
    relationship_type: Box<str>,
    target: Box<str>,
    mode: TargetMode,
}

impl SourceRelationship {
    fn capture(value: &Relationship, resource: &'static str) -> Result<Self> {
        Ok(Self {
            id: copy_boxed(value.r_id(), resource)?,
            relationship_type: copy_boxed(value.reltype(), resource)?,
            target: copy_boxed(value.target_ref(), resource)?,
            mode: value.target_mode(),
        })
    }

    fn matches(&self, value: &Relationship) -> bool {
        value.r_id() == self.id.as_ref()
            && value.reltype() == self.relationship_type.as_ref()
            && value.target_ref() == self.target.as_ref()
            && value.target_mode() == self.mode
    }
}

#[derive(Clone, Debug)]
struct SharedStringsState {
    relationship: SourceRelationship,
    part: PartState,
}

impl SharedStringsState {
    fn new(
        relationship: &Relationship,
        uri: PackURI,
        content_type: &str,
        bytes: SourcePayload,
    ) -> Result<Self> {
        Ok(Self {
            relationship: SourceRelationship::capture(
                relationship,
                "source-backed merge shared-strings relationship",
            )?,
            part: PartState::new(
                uri,
                content_type,
                bytes,
                "source-backed merge shared-strings content type",
            )?,
        })
    }

    fn same_source(&self, other: &Self) -> bool {
        self.relationship == other.relationship && self.part.same_source(&other.part)
    }

    fn matches_source_backed(
        &self,
        package: &SourceBackedPackage,
        workbook: &PartView<'_>,
    ) -> Result<bool> {
        let Some(relationship) = workbook.rels().get(self.relationship.id.as_ref()) else {
            return Ok(false);
        };
        if !self.relationship.matches(relationship) {
            return Ok(false);
        }
        let uri = relationship.target_partname()?;
        let part = package.part(&uri)?;
        let data = part.data()?;
        Ok(self.part.matches_view(&part, data.as_bytes()))
    }

    fn matches_current_source(&self, package: &OpcPackage, workbook: &dyn Part) -> bool {
        let Some(relationship) = workbook.rels().get(self.relationship.id.as_ref()) else {
            return false;
        };
        let Ok(uri) = relationship.target_partname() else {
            return false;
        };
        let Ok(part) = package.get_part(&uri) else {
            return false;
        };
        self.relationship.matches(relationship) && self.part.matches_part(part)
    }
}

fn check_execution(context: Option<&ExecutionContext>) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    context.check().map_err(|error| {
        Error::Package(match error {
            ExecutionError::Cancelled => OpcError::Cancelled,
            error => OpcError::Execution(error),
        })
    })
}

fn ensure_merge_area(sheet: &str, range: Rect) -> Result<()> {
    if range.rows() == 1 && range.columns() == 1 {
        return Err(Error::MergeEditBlocked {
            sheet: sheet.to_owned(),
            range,
            reason: MergeEditBlock::SingleCell,
        });
    }
    Ok(())
}

fn plan_changes(before: &[Rect], staged: &[Rect]) -> Result<(Vec<Rect>, Vec<Rect>)> {
    let mut add = Vec::new();
    add.try_reserve(staged.len())
        .map_err(|source| allocation("source-backed merge additions", source))?;
    for range in staged {
        if !before.contains(range) {
            add.push(*range);
        }
    }
    let mut remove = Vec::new();
    remove
        .try_reserve(before.len())
        .map_err(|source| allocation("source-backed merge removals", source))?;
    for range in before {
        if !staged.contains(range) {
            remove.push(*range);
        }
    }
    Ok((add, remove))
}

#[derive(Debug, Default)]
struct StructuralScan {
    protected: bool,
    merge_payload: bool,
    formulas: Vec<FormulaMarker>,
}

#[derive(Debug)]
struct FormulaMarker {
    kind: Box<str>,
    index: Option<u32>,
    range: Option<Rect>,
    cell: Option<Rect>,
}

fn validate_structural_preconditions(
    source: &Snapshot,
    add: &[Rect],
    remove: &[Rect],
) -> Result<()> {
    let requested = add
        .first()
        .or_else(|| remove.first())
        .copied()
        .ok_or_else(|| invalid("merged-range structural preflight lost its requested range"))?;
    let scan = scan_merge_structure(source.source_xml())?;
    if scan.protected {
        return Err(Error::MergeEditBlocked {
            sheet: source.sheet_name.to_string(),
            range: requested,
            reason: MergeEditBlock::ProtectedSheet,
        });
    }
    if scan.merge_payload {
        return Err(Error::MergeEditBlocked {
            sheet: source.sheet_name.to_string(),
            range: requested,
            reason: MergeEditBlock::UnmodeledPayload,
        });
    }
    let formula_ranges = structural_formula_ranges(&scan.formulas);
    for range in add {
        if formula_ranges
            .iter()
            .copied()
            .any(|formula| merge::overlaps(formula, *range))
        {
            return Err(Error::MergeEditBlocked {
                sheet: source.sheet_name.to_string(),
                range: *range,
                reason: MergeEditBlock::GroupFormula,
            });
        }
        if let Some(existing) = source
            .ranges
            .iter()
            .copied()
            .filter(|existing| !remove.contains(existing))
            .find(|existing| merge::overlaps(*existing, *range))
        {
            return Err(Error::MergeEditBlocked {
                sheet: source.sheet_name.to_string(),
                range: *range,
                reason: MergeEditBlock::Overlap { existing },
            });
        }
    }
    for (index, range) in add.iter().copied().enumerate() {
        if let Some(existing) = add
            .iter()
            .copied()
            .skip(index + 1)
            .find(|existing| merge::overlaps(*existing, range))
        {
            return Err(Error::MergeEditBlocked {
                sheet: source.sheet_name.to_string(),
                range,
                reason: MergeEditBlock::Overlap { existing },
            });
        }
    }
    Ok(())
}

fn scan_merge_structure(content: &[u8]) -> Result<StructuralScan> {
    let mut reader = Reader::from_reader(content);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut merge_cells_depth = None;
    let mut merge_cell_depth = None;
    let mut current_cell = None;
    let mut scan = StructuralScan::default();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) => {
                scan_structural_element(
                    &element,
                    depth,
                    &mut merge_cells_depth,
                    &mut merge_cell_depth,
                    &mut current_cell,
                    &mut scan,
                    reader.decoder(),
                )?;
                depth = depth.saturating_add(1);
            },
            Ok(Event::Empty(element)) => {
                scan_structural_element(
                    &element,
                    depth,
                    &mut merge_cells_depth,
                    &mut merge_cell_depth,
                    &mut current_cell,
                    &mut scan,
                    reader.decoder(),
                )?;
                let element_depth = depth.saturating_add(1);
                if merge_cell_depth == Some(element_depth) {
                    merge_cell_depth = None;
                }
                if merge_cells_depth == Some(element_depth) {
                    merge_cells_depth = None;
                }
            },
            Ok(Event::End(element)) => {
                let name = element.name().local_name();
                if merge_cell_depth == Some(depth) {
                    merge_cell_depth = None;
                }
                if merge_cells_depth == Some(depth) {
                    merge_cells_depth = None;
                }
                depth = depth.saturating_sub(1);
                if name.as_ref() == b"c" {
                    current_cell = None;
                }
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => {
                return Err(invalid(format!(
                    "merged-range XML preflight failed: {error}"
                )));
            },
        }
        buffer.clear();
    }
    Ok(scan)
}

fn scan_structural_element(
    element: &BytesStart<'_>,
    depth: usize,
    merge_cells_depth: &mut Option<usize>,
    merge_cell_depth: &mut Option<usize>,
    current_cell: &mut Option<Rect>,
    scan: &mut StructuralScan,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    let name = element.name().local_name();
    if *merge_cells_depth == Some(depth)
        && name.as_ref() != b"mergeCell"
        && name.as_ref() != b"mergeCells"
    {
        scan.merge_payload = true;
    }
    if merge_cell_depth.is_some() {
        scan.merge_payload = true;
    }
    if name.as_ref() == b"mergeCells" {
        *merge_cells_depth = Some(depth.saturating_add(1));
    } else if name.as_ref() == b"mergeCell" {
        *merge_cell_depth = Some(depth.saturating_add(1));
    } else if name.as_ref() == b"sheetProtection" {
        if attribute_value(element, b"sheet", decoder)?
            .is_some_and(|value| matches!(value.as_str(), "1" | "true" | "TRUE" | "True"))
        {
            scan.protected = true;
        }
    } else if name.as_ref() == b"c" {
        *current_cell = attribute_value(element, b"r", decoder)?
            .map(|value| Rect::from_a1(&value))
            .transpose()?;
    } else if name.as_ref() == b"f" {
        let kind = attribute_value(element, b"t", decoder)?.unwrap_or_else(|| "normal".to_owned());
        if matches!(kind.as_str(), "shared" | "array" | "dataTable") {
            let index = attribute_value(element, b"si", decoder)?
                .map(|value| value.parse::<u32>())
                .transpose()
                .map_err(|error| invalid(format!("invalid grouped formula index: {error}")))?;
            let range = attribute_value(element, b"ref", decoder)?
                .map(|value| Rect::from_a1(&value))
                .transpose()?;
            scan.formulas.push(FormulaMarker {
                kind: kind.into_boxed_str(),
                index,
                range,
                cell: *current_cell,
            });
        }
    }
    Ok(())
}

fn structural_formula_ranges(formulas: &[FormulaMarker]) -> Vec<Rect> {
    let mut shared = std::collections::HashMap::new();
    for formula in formulas {
        if formula.kind.as_ref() == "shared"
            && let (Some(index), Some(range)) = (formula.index, formula.range)
        {
            shared.insert(index, range);
        }
    }
    formulas
        .iter()
        .filter_map(|formula| {
            formula
                .range
                .or_else(|| formula.index.and_then(|index| shared.get(&index).copied()))
                .or(formula.cell)
        })
        .collect()
}

fn attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(false) {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid worksheet attribute: {error}")))?;
        if attribute.key.local_name().as_ref() == name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| invalid(format!("invalid worksheet attribute value: {error}")));
        }
    }
    Ok(None)
}

fn validate_follower_content(source: &Snapshot, store: &Store, staged: &[Rect]) -> Result<()> {
    for range in staged {
        if source.ranges.contains(range) {
            continue;
        }
        if let Some((address, _)) = store
            .cells(*range)
            .find(|(address, cell)| *address != range.start() && !matches!(cell, &Cell::Empty))
        {
            return Err(Error::MergeEditBlocked {
                sheet: source.sheet_name.to_string(),
                range: *range,
                reason: MergeEditBlock::FollowerContent { address },
            });
        }
    }
    Ok(())
}

fn canonical_ranges(ranges: Vec<Rect>) -> Result<Vec<Rect>> {
    let index = merge::Index::new(ranges)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(index.as_slice().len())
        .map_err(|source| allocation("canonical source-backed merged ranges", source))?;
    output.extend_from_slice(index.as_slice());
    output.sort_unstable_by_key(canonical_range_key);
    Ok(output)
}

fn canonical_range_key(range: &Rect) -> (u32, u32, u32, u32) {
    let (end_row, end_column) = range.end();
    (
        range.start().row().get(),
        range.start().column().get(),
        end_row,
        end_column,
    )
}

fn parse_worksheet_store(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
    xml: &[u8],
) -> Result<(Store, Option<Arc<[Text]>>, Option<SharedStringsState>)> {
    let mut strings = None;
    let mut shared_strings_source = None;
    let mut attempted = false;
    let store = raw::worksheet::parse(xml, || {
        if !attempted {
            attempted = true;
            let (loaded, source) = load_shared_strings(package, workbook)?;
            strings = loaded;
            shared_strings_source = source;
        }
        Ok(strings.as_deref())
    })?;
    Ok((store, strings.map(Arc::from), shared_strings_source))
}

fn parse_owned_worksheet_store(
    package: &OpcPackage,
    workbook: &dyn Part,
    xml: &[u8],
) -> Result<(Store, Option<Arc<[Text]>>, Option<SharedStringsState>)> {
    let mut strings = None;
    let mut shared_strings_source = None;
    let mut attempted = false;
    let store = raw::worksheet::parse(xml, || {
        if !attempted {
            attempted = true;
            let (loaded, source) = load_shared_strings_owned(package, workbook)?;
            strings = loaded;
            shared_strings_source = source;
        }
        Ok(strings.as_deref())
    })?;
    Ok((store, strings.map(Arc::from), shared_strings_source))
}

fn parse_worksheet_store_bytes(xml: &[u8], strings: Option<&[Text]>) -> Result<Store> {
    raw::worksheet::parse(xml, || Ok(strings))
}

fn load_shared_strings(
    package: &SourceBackedPackage,
    workbook: &PartView<'_>,
) -> Result<(Option<Box<[Text]>>, Option<SharedStringsState>)> {
    let mut found = None;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::SHARED_STRINGS | rt::STRICT_SHARED_STRINGS
        )
    }) {
        if found.is_some() {
            return Err(invalid("workbook has multiple shared-string relationships"));
        }
        if relationship.target_mode() != TargetMode::Internal {
            return Err(invalid("shared-string relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.part(&uri)?;
        if part.content_type() != ct::SML_SHARED_STRINGS {
            return Err(invalid(format!(
                "shared-string part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_SHARED_STRINGS
            )));
        }
        let payload = SourcePayload::from_part_data(package, part.data()?)?;
        let parsed = raw::strings::parse(payload.as_bytes())?;
        let source = SharedStringsState::new(
            relationship,
            part.partname().clone(),
            part.content_type(),
            payload,
        )?;
        found = Some((parsed, source));
    }
    Ok(found.map_or((None, None), |(strings, source)| {
        (Some(strings), Some(source))
    }))
}

fn load_shared_strings_owned(
    package: &OpcPackage,
    workbook: &dyn Part,
) -> Result<(Option<Box<[Text]>>, Option<SharedStringsState>)> {
    let mut found = None;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::SHARED_STRINGS | rt::STRICT_SHARED_STRINGS
        )
    }) {
        if found.is_some() {
            return Err(invalid("workbook has multiple shared-string relationships"));
        }
        if relationship.target_mode() != TargetMode::Internal {
            return Err(invalid("shared-string relationship cannot be external"));
        }
        let uri = relationship.target_partname()?;
        let part = package.get_part(&uri)?;
        if part.content_type() != ct::SML_SHARED_STRINGS {
            return Err(invalid(format!(
                "shared-string part has content type '{}', expected '{}'",
                part.content_type(),
                ct::SML_SHARED_STRINGS
            )));
        }
        let payload = SourcePayload::Owned(part.blob_arc());
        let parsed = raw::strings::parse(payload.as_bytes())?;
        let source = SharedStringsState::new(
            relationship,
            part.partname().clone(),
            part.content_type(),
            payload,
        )?;
        found = Some((parsed, source));
    }
    Ok(found.map_or((None, None), |(strings, source)| {
        (Some(strings), Some(source))
    }))
}

fn resolve_selector(sheets: &[raw::Sheet], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        CoreSelector::Position(position) => {
            Ok((position.get() < sheets.len()).then_some(position.get()))
        },
        CoreSelector::Name(name) => {
            let key = crate::sheet::key(&name);
            Ok(sheets
                .iter()
                .position(|sheet| crate::sheet::key(&sheet.name) == key))
        },
        CoreSelector::Id(never) => match never {},
        _ => Err(Error::UnsupportedSelector),
    }
}

fn require_workbook_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_SHEET_MACRO_MAIN
            | ct::SML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main part has non-XLSX content type '{content_type}'"
        )))
    }
}

fn require_selected_worksheet<'a>(
    relationships: &'a Relationships,
    sheet: &raw::Sheet,
    worksheet_uri: &PackURI,
    content_type: &str,
) -> Result<&'a Relationship> {
    let relationship = relationships
        .get(&sheet.relationship_id)
        .ok_or_else(|| invalid("selected worksheet relationship is missing"))?;
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET)
        || relationship.target_partname()? != *worksheet_uri
        || content_type != ct::SML_WORKSHEET
    {
        return Err(invalid(
            "selected worksheet relationship or content type is invalid",
        ));
    }
    Ok(relationship)
}

fn current_owner_relationship(relationships: &Relationships) -> Option<&Relationship> {
    let mut owners = relationships.iter().filter(|value| {
        matches!(
            value.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    });
    let owner = owners.next()?;
    if owners.next().is_some() || owner.target_mode() != TargetMode::Internal {
        return None;
    }
    Some(owner)
}

fn capture_relationships(relationships: &Relationships) -> Result<Vec<SourceRelationship>> {
    let mut captured = Vec::new();
    captured
        .try_reserve_exact(relationships.len())
        .map_err(|source| allocation("source-backed merge worksheet relationships", source))?;
    for relationship in relationships.iter() {
        captured.push(SourceRelationship::capture(
            relationship,
            "source-backed merge worksheet relationship",
        )?);
    }
    captured.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    Ok(captured)
}

fn relationships_match(expected: &[SourceRelationship], actual: &Relationships) -> bool {
    expected.len() == actual.len()
        && expected.iter().all(|value| {
            actual
                .get(value.id.as_ref())
                .is_some_and(|actual| value.matches(actual))
        })
}

fn optional_same<T: PartialEq>(left: &Option<T>, right: &Option<T>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => left == right,
        (None, _) | (_, None) => true,
    }
}

fn optional_shared_strings_same(
    left: &Option<SharedStringsState>,
    right: &Option<SharedStringsState>,
) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => left.same_source(right),
        (None, _) | (_, None) => true,
    }
}

fn copy_boxed(value: &str, resource: &'static str) -> Result<Box<str>> {
    let mut copied = String::new();
    copied
        .try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    copied.push_str(value);
    Ok(copied.into_boxed_str())
}
