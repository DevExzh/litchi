//! Exact-source, selector-first Keynote slide deletion.

use std::sync::Arc;
use std::{fmt, mem::size_of};

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{
    WireLimits,
    wire::{WireView, append_length_delimited_field_with_limits},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::package_metadata_codec::{
    ComponentDescriptor, ComponentSelector, DataReferenceOwnerRemoval, ExternalReferenceDescriptor,
    ExternalReferenceRemoval, ObjectUuidDescriptor, ObjectUuidRemoval, PackageMetadataVisitor,
    RemovalBatch, RewriteError, RewriteOptions, UuidBits, inspect_package_metadata_with_visitor,
    remove_package_metadata,
};
use thiserror::Error as ThisError;

use super::{
    Package, PhysicalSource, ReadError, SHOW_MESSAGE_TYPE, SLIDE_MESSAGE_TYPE,
    SLIDE_NODE_MESSAGE_TYPE, decode_show_snapshot, unique_payload,
};
use crate::{SlideSelector, SlideSelectorError};

const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

/// A content-free location associated with slide-deletion work.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Path {
    Package,
    Show,
    Slide { position: Position },
    PackageMetadata,
}

impl fmt::Display for Path {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => f.write_str("package"),
            Self::Show => f.write_str("show"),
            Self::Slide { position } => write!(f, "slide {position:?}"),
            Self::PackageMetadata => f.write_str("package metadata"),
        }
    }
}

/// A finite resource governed by a slide-deletion transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    Objects,
    Slides,
    References,
    TextStorages,
    TextFragments,
    TextBytes,
    WireBytes,
    WireFields,
    WireNesting,
    Work,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Objects => "objects",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::Work => "work",
        })
    }
}

/// A slide-deletion failure that never exposes native identifiers or bytes.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    #[error("this Keynote source does not support physical edits")]
    UnsupportedSource,
    #[error("this Keynote slide topology is not supported for deletion")]
    UnsupportedTopology,
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the final Keynote slide cannot be deleted")]
    CannotDeleteFinalSlide,
    #[error("the Keynote slide-deletion transaction already has a staged operation")]
    OperationAlreadyStaged,
    #[error("the Keynote slide-deletion transaction has no staged operation")]
    NoStagedOperation,
    #[error("the Keynote source cannot be deleted from safely")]
    InvalidSource,
    #[error("the selected Keynote slide graph has another surviving owner")]
    AmbiguousOwnership,
    #[error(
        "Keynote slide-deletion {kind} limit exceeded at {path}: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
        path: Path,
    },
    #[error("could not allocate {amount} units for the Keynote slide-deletion transaction")]
    Allocation { amount: usize },
    #[error("the deleted Keynote candidate failed semantic verification")]
    Verification,
    #[error("the Keynote slide-deletion patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy)]
struct Intent {
    position: Position,
}

#[derive(Debug)]
struct Selection<'a> {
    position: Position,
    show_identifier: u64,
    node_identifier: u64,
    slide_identifier: u64,
    show_component: &'a str,
    node_component: &'a str,
    slide_component: &'a str,
    show_message_index: usize,
    source_slide_names: Box<[Option<Box<str>>]>,
    metadata: MetadataSelection<'a>,
}

#[derive(Debug)]
struct MetadataSelection<'a> {
    component: &'a str,
    object_identifier: u64,
    message_index: usize,
    rewritten_payload: Vec<u8>,
}

struct MetadataFacts<'a> {
    node_component: Option<u64>,
    slide_component: Option<u64>,
    node_uuid: Option<UuidBits>,
    slide_uuid: Option<UuidBits>,
    external_record: Option<(u64, Option<bool>)>,
    component_external_record: Option<(u64, Option<bool>)>,
    invalid: bool,
    ambiguous: bool,
    node_identifier: u64,
    slide_identifier: u64,
    node_locator: &'a str,
    slide_locator: &'a str,
}

impl PackageMetadataVisitor for MetadataFacts<'_> {
    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), RewriteError> {
        if component.is_current() && component.effective_locator() == self.node_locator {
            if self
                .node_component
                .replace(component.identifier())
                .is_some()
            {
                self.invalid = true;
            }
        }
        if component.is_current() && component.effective_locator() == self.slide_locator {
            if self
                .slide_component
                .replace(component.identifier())
                .is_some()
            {
                self.invalid = true;
            }
        }
        Ok(())
    }
    fn visit_object_uuid(&mut self, binding: ObjectUuidDescriptor<'_>) -> Result<(), RewriteError> {
        if binding.object_identifier() == self.node_identifier {
            if !binding.component().is_current()
                || binding.component().effective_locator() != self.node_locator
                || self.node_uuid.replace(binding.uuid()).is_some()
            {
                self.invalid = true;
            }
        }
        if binding.object_identifier() == self.slide_identifier {
            if !binding.component().is_current()
                || binding.component().effective_locator() != self.slide_locator
                || self.slide_uuid.replace(binding.uuid()).is_some()
            {
                self.invalid = true;
            }
        }
        Ok(())
    }
    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        if reference.object_identifier() == Some(self.node_identifier) {
            if reference.is_versioned() {
                self.invalid = true;
            } else {
                self.ambiguous = true;
            }
        }
        if reference.object_identifier() == Some(self.slide_identifier) {
            if reference.is_versioned() {
                self.invalid = true;
            } else if !reference.source().is_current()
                || reference.source().effective_locator() != self.node_locator
            {
                self.ambiguous = true;
            } else if self
                .external_record
                .replace((reference.target_component_identifier(), reference.is_weak()))
                .is_some()
            {
                self.invalid = true;
            }
        }
        if reference.object_identifier().is_none()
            && reference.target_component_identifier() == self.slide_identifier
        {
            if reference.is_versioned() {
                self.invalid = true;
            } else if !reference.source().is_current()
                || reference.source().effective_locator() != self.node_locator
            {
                self.ambiguous = true;
            } else if self
                .component_external_record
                .replace((reference.target_component_identifier(), reference.is_weak()))
                .is_some()
            {
                self.invalid = true;
            }
        }
        Ok(())
    }
}

fn normalized_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

fn metadata_options(
    source: &Package,
    budget: &Budget,
    removals: usize,
) -> Result<RewriteOptions, Error> {
    let wire = source.wire_limits().map_err(map_wire_error)?;
    let semantic = source.semantic_limits();
    Ok(RewriteOptions::new(
        wire.max_input_bytes(),
        wire.max_output_bytes(),
        wire.max_fields(),
        wire.max_rewrite_work().min(budget.remaining_work()),
        64,
        semantic.max_objects(),
        semantic.max_references().min(budget.remaining_references()),
        removals,
    ))
}

fn select_metadata<'a>(
    source: &'a Package,
    node_identifier: u64,
    slide_identifier: u64,
    node_component: &'a str,
    slide_component: &'a str,
    budget: &mut Budget,
) -> Result<MetadataSelection<'a>, Error> {
    let mut route = None;
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                    if route
                        .replace((component.name(), object, index, message.data.as_slice()))
                        .is_some()
                    {
                        return Err(Error::InvalidSource);
                    }
                }
            }
        }
    }
    let (metadata_component, object, message_index, payload) = route.ok_or(Error::InvalidSource)?;
    let mut facts = MetadataFacts {
        node_component: None,
        slide_component: None,
        node_uuid: None,
        slide_uuid: None,
        external_record: None,
        component_external_record: None,
        invalid: false,
        ambiguous: false,
        node_identifier,
        slide_identifier,
        node_locator: normalized_locator(node_component),
        slide_locator: normalized_locator(slide_component),
    };
    let inspection = inspect_package_metadata_with_visitor(
        payload,
        metadata_options(source, budget, 0)?,
        &mut facts,
    )
    .map_err(map_metadata_error)?;
    budget.merge_metadata_report(inspection.report())?;
    let node_component_id = facts.node_component.ok_or(Error::InvalidSource)?;
    let slide_component_id = facts.slide_component.ok_or(Error::InvalidSource)?;
    if facts
        .external_record
        .is_some_and(|(target, _)| target != slide_component_id)
        || facts
            .component_external_record
            .is_some_and(|(target, _)| target != slide_component_id)
    {
        return Err(Error::InvalidSource);
    }
    if facts.external_record.is_some() == facts.component_external_record.is_some() {
        return Err(Error::InvalidSource);
    }
    if facts.invalid {
        return Err(Error::InvalidSource);
    }
    if facts.ambiguous {
        return Err(Error::AmbiguousOwnership);
    }
    let node_selector = ComponentSelector::new(node_component_id, facts.node_locator);
    let slide_selector = ComponentSelector::new(slide_component_id, facts.slide_locator);
    let node_info = source
        .object(node_identifier)
        .ok_or(Error::InvalidSource)?
        .archive_info
        .message_infos
        .first()
        .ok_or(Error::InvalidSource)?;
    let slide_info = source
        .object(slide_identifier)
        .ok_or(Error::InvalidSource)?
        .archive_info
        .message_infos
        .first()
        .ok_or(Error::InvalidSource)?;
    let data_count = node_info
        .data_references
        .len()
        .checked_add(slide_info.data_references.len())
        .ok_or(Error::InvalidSource)?;
    let mut data_owners = Vec::new();
    data_owners
        .try_reserve_exact(data_count)
        .map_err(|_| Error::Allocation { amount: data_count })?;
    data_owners.extend(
        node_info
            .data_references
            .iter()
            .map(|id| DataReferenceOwnerRemoval::new(node_selector, *id, node_identifier, 1)),
    );
    data_owners.extend(
        slide_info
            .data_references
            .iter()
            .map(|id| DataReferenceOwnerRemoval::new(slide_selector, *id, slide_identifier, 1)),
    );
    let node_uuid = facts.node_uuid.ok_or(Error::InvalidSource)?;
    let slide_uuid = facts.slide_uuid.ok_or(Error::InvalidSource)?;
    let uuid_removals = [
        ObjectUuidRemoval::new(node_selector, node_identifier, node_uuid),
        ObjectUuidRemoval::new(slide_selector, slide_identifier, slide_uuid),
    ];
    // Native Keynote packages may register only the retained Slide component,
    // with no object-specific external reference for its slide object. Remove
    // an exact object-specific record when present; a component-level record
    // (`object_identifier == None`) is not object ownership and is preserved.
    let external_removal = facts.external_record.map(|(_, weakness)| {
        ExternalReferenceRemoval::new(node_selector, slide_selector, slide_identifier, weakness)
    });
    let external_removals = external_removal.as_slice();
    let removal = remove_package_metadata(
        payload,
        RemovalBatch::new(
            inspection.last_object_identifier(),
            &uuid_removals,
            external_removals,
            &data_owners,
        ),
        metadata_options(
            source,
            budget,
            uuid_removals.len() + external_removals.len() + data_owners.len(),
        )?,
    )
    .map_err(map_metadata_error)?;
    budget.merge_metadata_report(removal.report())?;
    let rewritten_payload = removal.into_bytes();
    Ok(MetadataSelection {
        component: metadata_component,
        object_identifier: object.archive_info.identifier.ok_or(Error::InvalidSource)?,
        message_index,
        rewritten_payload,
    })
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct TestReport {
    work: u64,
    objects_scanned: u64,
    references_scanned: u64,
    allocation_events: u64,
    peak_scratch_bytes: u64,
    component_deletions: u64,
    output_allocations: u64,
    candidate_reopens: u64,
}

#[cfg(test)]
pub(super) fn production_test_attempt(
    package: &Package,
    position: Position,
    max_work: Option<u64>,
) -> (Result<TestReport, Error>, TestReport) {
    let mut budget = match Budget::for_package(package, None) {
        Ok(budget) => budget,
        Err(error) => return (Err(error), TestReport::default()),
    };
    let prepared = prepare_selection(package, Intent { position }, &mut budget);
    // Record the exact full preparation consumption before enforcing the test
    // ceiling so failures retain auditable no-publication counters.
    let observed = budget.report.work;
    if let Some(maximum) = max_work {
        if observed > maximum {
            let error = Error::LimitExceeded {
                kind: LimitKind::Work,
                observed,
                maximum,
                path: Path::Package,
            };
            return (Err(error), budget.report);
        }
    }
    let result = prepared
        .and_then(|selection| publish_selection(package, selection, &mut budget))
        .map(|_| budget.report);
    (result, budget.report)
}

#[cfg(test)]
impl TestReport {
    pub(super) const fn work(self) -> u64 {
        self.work
    }
    pub(super) const fn objects_scanned(self) -> u64 {
        self.objects_scanned
    }
    pub(super) const fn references_scanned(self) -> u64 {
        self.references_scanned
    }
    pub(super) const fn allocation_events(self) -> u64 {
        self.allocation_events
    }
    pub(super) const fn peak_scratch_bytes(self) -> u64 {
        self.peak_scratch_bytes
    }
    pub(super) const fn component_deletions(self) -> u64 {
        self.component_deletions
    }
    pub(super) const fn output_allocations(self) -> u64 {
        self.output_allocations
    }
    pub(super) const fn candidate_reopens(self) -> u64 {
        self.candidate_reopens
    }
}

struct Budget {
    maximum: u64,
    max_objects: u64,
    max_references: u64,
    report: TestReport,
}

impl Budget {
    fn for_package(package: &Package, maximum: Option<u64>) -> Result<Self, Error> {
        let semantic = package.semantic_limits();
        let wire_maximum = package
            .wire_limits()
            .map_err(map_wire_error)?
            .max_rewrite_work() as u64;
        Ok(Self {
            maximum: maximum.map_or(wire_maximum, |limit| limit.min(wire_maximum)),
            max_objects: semantic.max_objects() as u64,
            max_references: semantic.max_references() as u64,
            report: TestReport::default(),
        })
    }
    fn charge_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        self.report.work = self
            .report
            .work
            .checked_add(amount as u64)
            .ok_or(Error::InvalidSource)?;
        if self.report.work > self.maximum {
            return Err(Error::LimitExceeded {
                kind: LimitKind::Work,
                observed: self.report.work,
                maximum: self.maximum,
                path,
            });
        }
        Ok(())
    }
    fn charge_reference(&mut self) -> Result<(), Error> {
        self.report.references_scanned = self
            .report
            .references_scanned
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        if self.report.references_scanned > self.max_references {
            return Err(Error::LimitExceeded {
                kind: LimitKind::References,
                observed: self.report.references_scanned,
                maximum: self.max_references,
                path: Path::Package,
            });
        }
        self.charge_work(1, Path::Package)
    }
    fn charge_object(&mut self) -> Result<(), Error> {
        self.report.objects_scanned = self
            .report
            .objects_scanned
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        if self.report.objects_scanned > self.max_objects {
            return Err(Error::LimitExceeded {
                kind: LimitKind::Objects,
                observed: self.report.objects_scanned,
                maximum: self.max_objects,
                path: Path::Package,
            });
        }
        self.charge_work(1, Path::Package)
    }
    fn allocation(&mut self, scratch_bytes: usize) -> Result<(), Error> {
        self.report.allocation_events = self
            .report
            .allocation_events
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        self.report.peak_scratch_bytes = self.report.peak_scratch_bytes.max(scratch_bytes as u64);
        Ok(())
    }
    fn merge_metadata_report(
        &mut self,
        report: litchi_iwa_protos::package_metadata_codec::RewriteReport,
    ) -> Result<(), Error> {
        self.charge_work(report.work_bytes(), Path::PackageMetadata)?;
        self.report.references_scanned = self
            .report
            .references_scanned
            .checked_add(report.references_scanned() as u64)
            .ok_or(Error::InvalidSource)?;
        if self.report.references_scanned > self.max_references {
            return Err(Error::LimitExceeded {
                kind: LimitKind::References,
                observed: self.report.references_scanned,
                maximum: self.max_references,
                path: Path::PackageMetadata,
            });
        }
        self.report.allocation_events = self
            .report
            .allocation_events
            .checked_add(report.allocations() as u64)
            .ok_or(Error::InvalidSource)?;
        self.report.peak_scratch_bytes = self
            .report
            .peak_scratch_bytes
            .max(report.scratch_bytes() as u64);
        Ok(())
    }
    const fn remaining_work(&self) -> usize {
        self.maximum.saturating_sub(self.report.work) as usize
    }
    const fn remaining_references(&self) -> usize {
        self.max_references
            .saturating_sub(self.report.references_scanned) as usize
    }
}

/// One bounded slide-deletion edit staged against an immutable package.
#[derive(Debug)]
pub struct Edit<'a> {
    source: &'a Package,
    intent: Option<Intent>,
}

impl<'a> Edit<'a> {
    pub(super) const fn new(source: &'a Package) -> Self {
        Self {
            source,
            intent: None,
        }
    }

    /// Stage removal of one semantic slide.
    pub fn remove_slide<'s>(
        &mut self,
        selector: impl Into<SlideSelector<'s>>,
    ) -> Result<&mut Self, Error> {
        if self.intent.is_some() {
            return Err(Error::OperationAlreadyStaged);
        }
        let selector = selector.into();
        let show = self.source.show().map_err(map_read_error)?;
        if show.slides().len() <= 1 {
            return Err(Error::CannotDeleteFinalSlide);
        }
        let selected = show
            .select_slide(selector)
            .map_err(map_selector_error)?
            .ok_or(match selector {
                SlideSelector::Name(_) => Error::SlideNameNotFound,
                SlideSelector::Position(position) => Error::SlidePositionNotFound { position },
            })?;
        self.intent = Some(Intent {
            position: Position::new(selected.index()),
        });
        Ok(self)
    }

    /// Validate, rewrite, reopen, and publish the staged immutable candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let intent = self.intent.ok_or(Error::NoStagedOperation)?;
        commit_deletion(self.source, intent)
    }
}

/// Exact-source-checked reversible slide-deletion patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    position: Position,
    source_slide_count: usize,
    target_slide_count: usize,
    source_slide_names: Arc<[Option<Box<str>>]>,
    target_slide_names: Arc<[Option<Box<str>>]>,
    touched_components: usize,
    removes_objects: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl Patch {
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            position: self.position,
            source_slide_count: self.target_slide_count,
            target_slide_count: self.source_slide_count,
            source_slide_names: Arc::clone(&self.target_slide_names),
            target_slide_names: Arc::clone(&self.source_slide_names),
            touched_components: self.touched_components,
            removes_objects: !self.removes_objects,
        }
    }
}

/// Compact evidence for one verified deletion.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    slides_removed: usize,
    slides_restored: usize,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    #[must_use]
    pub const fn changed(self) -> bool {
        true
    }
    #[must_use]
    pub const fn slides_removed(self) -> usize {
        self.slides_removed
    }
    #[must_use]
    pub const fn slides_restored(self) -> usize {
        self.slides_restored
    }
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The verified result of one immutable slide deletion.
#[must_use]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Start one selector-first slide deletion.
    #[must_use]
    pub const fn edit_slide_deletion(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Apply a deletion patch only to its exact immutable source artifact.
    pub fn apply_slide_deletion(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let source_show = self.show().map_err(map_read_error)?;
        if source_show.slides().len() != patch.source_slide_count {
            return Err(Error::PatchConflict);
        }
        if !slide_names_match(source_show, &patch.source_slide_names) {
            return Err(Error::PatchConflict);
        }
        let target = patch.artifacts.target();
        let candidate = Package::from_source_with_options(Arc::clone(&target), self.state.options)
            .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let candidate_show = candidate.show().map_err(map_read_error)?;
        if candidate_show.slides().len() != patch.target_slide_count
            || !slide_names_match(candidate_show, &patch.target_slide_names)
        {
            return Err(Error::PatchConflict);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                slides_removed: usize::from(patch.removes_objects),
                slides_restored: usize::from(!patch.removes_objects),
                touched_components: patch.touched_components,
                full_reparse_performed: true,
            },
        })
    }
}

fn commit_deletion(source: &Package, intent: Intent) -> Result<Commit, Error> {
    source.validate().map_err(map_read_error)?;
    let source_bytes = physical_source(source)?;
    let source_slide_count = source.show().map_err(map_read_error)?.slides().len();
    let package = rewrite_deletion(source, intent)?;
    let target = physical_source(&package)?;
    let touched_components = count_changed_components(source, &package);
    let source_slide_names = semantic_slide_names(source)?;
    let target_slide_names = semantic_slide_names(&package)?;
    Ok(Commit {
        package,
        patch: Patch {
            artifacts: ExactArtifacts::new(source_bytes, target),
            position: intent.position,
            source_slide_count,
            target_slide_count: source_slide_count
                .checked_sub(1)
                .ok_or(Error::InvalidSource)?,
            source_slide_names,
            target_slide_names,
            touched_components,
            removes_objects: true,
        },
        diagnostics: Diagnostics {
            slides_removed: 1,
            slides_restored: 0,
            touched_components,
            full_reparse_performed: true,
        },
    })
}

fn semantic_slide_names(package: &Package) -> Result<Arc<[Option<Box<str>>]>, Error> {
    let show = package.show().map_err(map_read_error)?;
    Ok(show
        .slides()
        .iter()
        .map(|slide| slide.name().map(Into::into))
        .collect::<Vec<_>>()
        .into())
}

fn slide_names_match(show: &crate::show::Show, expected: &[Option<Box<str>>]) -> bool {
    show.slides().len() == expected.len()
        && show
            .slides()
            .iter()
            .zip(expected)
            .all(|(slide, expected)| slide.name() == expected.as_deref())
}

fn rewrite_deletion(source: &Package, intent: Intent) -> Result<Package, Error> {
    let mut budget = Budget::for_package(source, None)?;
    let selection = prepare_selection(source, intent, &mut budget)?;
    let target = publish_selection(source, selection, &mut budget)?;
    Ok(target)
}

fn prepare_selection<'a>(
    source: &'a Package,
    intent: Intent,
    budget: &mut Budget,
) -> Result<Selection<'a>, Error> {
    let show_identifier = source.root_show_identifier().map_err(map_read_error)?;
    let (show_component, show) = source
        .object_with_component(show_identifier)
        .ok_or(Error::InvalidSource)?;
    let (show_message_index, show_payload) = one_message(show, SHOW_MESSAGE_TYPE)?;
    let limits = source.wire_limits().map_err(map_wire_error)?;
    let snapshot =
        decode_show_snapshot(show_payload, source.semantic_limits().max_slides(), limits)
            .map_err(map_read_error)?;
    let nodes = snapshot.slide_node_identifiers();
    if nodes.len() <= 1 {
        return Err(Error::CannotDeleteFinalSlide);
    }
    if intent.position.get() >= nodes.len() {
        return Err(Error::InvalidSource);
    }
    if snapshot.has_deprecated_root_slide_node() || snapshot.has_slide_list() {
        return Err(Error::UnsupportedTopology);
    }
    budget.allocation(
        nodes
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut sorted = nodes.to_vec();
    sorted.sort_unstable();
    if sorted.first() == Some(&0) || sorted.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Error::UnsupportedTopology);
    }
    let node_identifier = nodes[intent.position.get()];
    validate_selected_reference_metadata(show, show_message_index, node_identifier, &[3, 2])?;
    let (node_component, node) = source
        .object_with_component(node_identifier)
        .ok_or(Error::InvalidSource)?;
    let (node_message_index, node_payload) = one_message(node, SLIDE_NODE_MESSAGE_TYPE)?;
    if node.messages.len() != 1
        || node.archive_info.message_infos.len() != 1
        || node_message_index != 0
    {
        return Err(Error::InvalidSource);
    }
    let slide_identifier = strict_flat_slide(node_payload, limits)?;
    validate_selected_reference_metadata(node, node_message_index, slide_identifier, &[2])?;
    let (slide_component, slide) = source
        .object_with_component(slide_identifier)
        .ok_or(Error::InvalidSource)?;
    let _ = unique_payload(&slide.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
        .map_err(map_read_error)?;
    if slide.messages.len() != 1 || slide.archive_info.message_infos.len() != 1 {
        return Err(Error::InvalidSource);
    }
    validate_data_reference_metadata(&node.archive_info.message_infos[0])?;
    validate_data_reference_metadata(&slide.archive_info.message_infos[0])?;
    let metadata = select_metadata(
        source,
        node_identifier,
        slide_identifier,
        node_component,
        slide_component,
        budget,
    )?;

    let mut node_owner = 0usize;
    let mut slide_owner = 0usize;
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            budget.charge_object()?;
            let owner = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            for (message_index, info) in object.archive_info.message_infos.iter().enumerate() {
                if owner == slide_identifier && message_index == 0 {
                    validate_reference_metadata_consistency(info)?;
                }
                if object.archive_info.should_merge == Some(true)
                    || info.base_message_index.is_some()
                    || !info.diff_merge_version.is_empty()
                    || info.diff_field_path.is_some()
                    || !info.fields_to_remove.is_empty()
                    || !info.diff_read_version.is_empty()
                {
                    return Err(Error::InvalidSource);
                }
                budget.charge_work(info.field_infos.len().saturating_add(1), Path::Package)?;
                for id in &info.object_references {
                    budget.charge_reference()?;
                    if *id == node_identifier {
                        node_owner += 1;
                        if owner != show_identifier || message_index != show_message_index {
                            return Err(Error::AmbiguousOwnership);
                        }
                    }
                    if *id == slide_identifier {
                        slide_owner += 1;
                        if owner != node_identifier || message_index != node_message_index {
                            return Err(Error::AmbiguousOwnership);
                        }
                    }
                }
                for id in &info.data_references {
                    budget.charge_reference()?;
                    if *id == node_identifier || *id == slide_identifier {
                        return Err(Error::InvalidSource);
                    }
                }
                for field in &info.field_infos {
                    budget.charge_work(field.path.as_slice().len(), Path::Package)?;
                    for id in &field.object_references {
                        budget.charge_reference()?;
                        if *id == node_identifier
                            && (owner != show_identifier
                                || message_index != show_message_index
                                || field.path.as_slice() != [3, 2])
                        {
                            return Err(Error::AmbiguousOwnership);
                        }
                        if *id == slide_identifier
                            && (owner != node_identifier
                                || message_index != node_message_index
                                || field.path.as_slice() != [2])
                        {
                            return Err(Error::AmbiguousOwnership);
                        }
                    }
                    for id in &field.data_references {
                        budget.charge_reference()?;
                        if *id == node_identifier || *id == slide_identifier {
                            return Err(Error::InvalidSource);
                        }
                    }
                }
            }
        }
    }
    if node_owner != 1 || slide_owner != 1 {
        return Err(Error::InvalidSource);
    }
    let semantic_show = source.show().map_err(map_read_error)?;
    budget.allocation(
        semantic_show
            .slides()
            .len()
            .checked_mul(size_of::<Option<Box<str>>>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let source_slide_names = semantic_show
        .slides()
        .iter()
        .map(|slide| slide.name().map(Into::into))
        .collect::<Vec<_>>()
        .into_boxed_slice();
    Ok(Selection {
        position: intent.position,
        show_identifier,
        node_identifier,
        slide_identifier,
        show_component,
        node_component,
        slide_component,
        show_message_index,
        source_slide_names,
        metadata,
    })
}

fn validate_reference_metadata_consistency(
    info: &litchi_iwa_core::MessageInfo,
) -> Result<(), Error> {
    let mut aggregate = info.object_references.clone();
    aggregate.sort_unstable();
    if aggregate.first() == Some(&0) || aggregate.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Error::InvalidSource);
    }
    let mut attributed = Vec::new();
    for field in &info.field_infos {
        if !field.object_references.is_empty() && field.r#type != Some(FieldType::ObjectReference) {
            return Err(Error::InvalidSource);
        }
        attributed.extend_from_slice(&field.object_references);
    }
    attributed.sort_unstable();
    if attributed.windows(2).any(|pair| pair[0] == pair[1])
        || attributed
            .iter()
            .any(|id| aggregate.binary_search(id).is_err())
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_data_reference_metadata(info: &litchi_iwa_core::MessageInfo) -> Result<(), Error> {
    let mut aggregate = Vec::new();
    aggregate
        .try_reserve_exact(info.data_references.len())
        .map_err(|_| Error::Allocation {
            amount: info.data_references.len(),
        })?;
    aggregate.extend_from_slice(&info.data_references);
    aggregate.sort_unstable();
    if aggregate.first() == Some(&0) || aggregate.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(Error::InvalidSource);
    }
    let mut attributed = Vec::new();
    attributed
        .try_reserve_exact(info.data_references.len())
        .map_err(|_| Error::Allocation {
            amount: info.data_references.len(),
        })?;
    for field in &info.field_infos {
        if !field.data_references.is_empty() {
            if field.r#type != Some(FieldType::DataReference) {
                return Err(Error::InvalidSource);
            }
            attributed.extend_from_slice(&field.data_references);
        }
    }
    attributed.sort_unstable();
    if !attributed.is_empty() && attributed != aggregate {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn one_message(object: &ArchiveObject, type_: u32) -> Result<(usize, &[u8]), Error> {
    let mut matches = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, m)| m.type_ == type_);
    let (index, message) = matches.next().ok_or(Error::InvalidSource)?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok((index, &message.data))
}

fn strict_flat_slide(payload: &[u8], limits: WireLimits) -> Result<u64, Error> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut selected = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 => return Err(Error::UnsupportedTopology),
            2 => {
                if selected.is_some() || field.wire_type() != 2 {
                    return Err(Error::UnsupportedTopology);
                }
                selected = Some(
                    super::validate_reference_payload(
                        field.payload(),
                        limits,
                        "Keynote slide reference",
                    )
                    .map_err(map_wire_error)?,
                );
            },
            21 => {
                if field.wire_type() != 0 || strict_canonical_varint(field.payload())? != 1 {
                    return Err(Error::UnsupportedTopology);
                }
            },
            _ => {},
        }
    }
    selected
        .filter(|id| *id != 0)
        .ok_or(Error::UnsupportedTopology)
}

fn strict_canonical_varint(payload: &[u8]) -> Result<u64, Error> {
    let mut value = 0u64;
    for (index, byte) in payload.iter().copied().enumerate() {
        if index >= 10 || (index == 9 && byte > 1) {
            return Err(Error::InvalidSource);
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let canonical_len = if value == 0 {
                1
            } else {
                ((64 - value.leading_zeros()) as usize).div_ceil(7)
            };
            return (index + 1 == payload.len() && payload.len() == canonical_len)
                .then_some(value)
                .ok_or(Error::InvalidSource);
        }
    }
    Err(Error::InvalidSource)
}

fn validate_selected_reference_metadata(
    object: &ArchiveObject,
    message_index: usize,
    identifier: u64,
    path: &[u32],
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if info
        .object_references
        .iter()
        .filter(|id| **id == identifier)
        .count()
        != 1
        || info.data_references.contains(&identifier)
    {
        return Err(Error::InvalidSource);
    }
    let mut attributed = false;
    for field in &info.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|id| **id == identifier)
            .count();
        if count != 0 {
            if count != 1
                || attributed
                || field.path.as_slice() != path
                || field.r#type != Some(FieldType::ObjectReference)
            {
                return Err(Error::InvalidSource);
            }
            attributed = true;
        }
        if field.data_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
    }
    // Native producers may omit FieldInfo for an otherwise exact aggregate
    // reference. When present it is validated above; the strict payload path
    // and unique aggregate occurrence remain authoritative.
    Ok(())
}

fn remove_show_node(payload: &[u8], identifier: u64, limits: WireLimits) -> Result<Vec<u8>, Error> {
    let show = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut found_tree = false;
    let mut output = Vec::new();
    output
        .try_reserve_exact(payload.len())
        .map_err(|_| Error::Allocation {
            amount: payload.len(),
        })?;
    for field in show.fields() {
        if field.number() != 3 {
            output.extend_from_slice(field.raw());
            continue;
        }
        if found_tree || field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        found_tree = true;
        let tree = WireView::parse_with_limits(field.payload(), limits).map_err(map_wire_error)?;
        let mut retained = Vec::new();
        let mut removed = 0usize;
        for child in tree.fields() {
            if child.number() == 2 {
                if child.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                let id = super::validate_reference_payload(
                    child.payload(),
                    limits,
                    "Keynote slide-node reference",
                )
                .map_err(map_wire_error)?;
                if id == identifier {
                    removed += 1;
                    continue;
                }
            }
            retained
                .try_reserve(child.raw().len())
                .map_err(|_| Error::Allocation {
                    amount: child.raw().len(),
                })?;
            retained.extend_from_slice(child.raw());
        }
        if removed != 1 {
            return Err(Error::InvalidSource);
        }
        append_length_delimited_field_with_limits(&mut output, 3, &retained, limits)
            .map_err(map_wire_error)?;
    }
    if !found_tree {
        return Err(Error::InvalidSource);
    }
    Ok(output)
}

fn publish_selection(
    source: &Package,
    selection: Selection<'_>,
    budget: &mut Budget,
) -> Result<Package, Error> {
    let catalog = match &source.state.source {
        PhysicalSource::Package(catalog) if catalog.source_is_exact() => catalog,
        PhysicalSource::Package(_) => return Err(Error::UnsupportedSource),
        PhysicalSource::Semantic(_) => return Err(Error::UnsupportedSource),
    };
    let limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let wire_limits = source.wire_limits().map_err(map_wire_error)?;
    let names = [
        selection.show_component,
        selection.node_component,
        selection.slide_component,
        selection.metadata.component,
    ];
    budget.allocation(names.len().saturating_mul(size_of::<(&str, Vec<u8>)>()))?;
    let mut rewritten: Vec<(&str, Vec<u8>)> = Vec::new();
    for (index, name) in names.iter().enumerate() {
        if names[..index].contains(name) {
            continue;
        }
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == *name)
            .ok_or(Error::InvalidSource)?;
        if entry.is_opaque() {
            return Err(Error::InvalidSource);
        }
        let stream = SnappyStream::decompress_with_limits(
            entry.data(),
            catalog
                .limits()
                .snappy_limits()
                .map_err(map_archive_error)?,
        )
        .map_err(map_core_error)?;
        let mut archive =
            Archive::parse_with_limits(stream.as_bytes(), limits).map_err(map_core_error)?;
        if *name == selection.show_component {
            let show = archive
                .object(selection.show_identifier)
                .ok_or(Error::InvalidSource)?;
            let message = show
                .messages
                .get(selection.show_message_index)
                .ok_or(Error::InvalidSource)?;
            let payload = remove_show_node(&message.data, selection.node_identifier, wire_limits)?;
            archive
                .object_mut(selection.show_identifier)
                .ok_or(Error::InvalidSource)?
                .replace_message_pruning_object_references_preserving_header_with_limits(
                    selection.show_message_index,
                    RawMessage {
                        type_: SHOW_MESSAGE_TYPE,
                        data: payload,
                    },
                    &[selection.node_identifier],
                    limits,
                )
                .map_err(map_core_error)?;
        }
        if *name == selection.node_component
            && archive.remove_object(selection.node_identifier).is_none()
        {
            return Err(Error::InvalidSource);
        }
        if *name == selection.slide_component
            && archive.remove_object(selection.slide_identifier).is_none()
        {
            return Err(Error::InvalidSource);
        }
        if *name == selection.metadata.component {
            archive
                .object_mut(selection.metadata.object_identifier)
                .ok_or(Error::InvalidSource)?
                .replace_message_preserving_header_with_limits(
                    selection.metadata.message_index,
                    RawMessage {
                        type_: PACKAGE_METADATA_MESSAGE_TYPE,
                        data: selection.metadata.rewritten_payload.clone(),
                    },
                    limits,
                )
                .map_err(map_core_error)?;
        }
        verify_expected_archive_delta(
            component_archive(source, name)?,
            &archive,
            name,
            &selection,
            limits,
            wire_limits,
        )?;
        let bytes = archive
            .to_bytes_with_limits(limits)
            .map_err(map_core_error)?;
        let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
        budget.allocation(compressed.len())?;
        rewritten.push((name, compressed));
    }
    let edits: Vec<_> = rewritten
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect();
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| Error::InvalidSource)?;
    budget.report.output_allocations += 1;
    let output = catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, previews.names(), catalog.limits())
        .map_err(map_archive_error)?;
    budget.report.candidate_reopens += 1;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    if !super::rendering_invalidation::root_previews_absent(candidate.state.source.package())
        .map_err(|_| Error::Verification)?
    {
        return Err(Error::Verification);
    }
    let before = source.show().map_err(map_read_error)?;
    let after = candidate.show().map_err(map_read_error)?;
    let expected_names = selection
        .source_slide_names
        .iter()
        .enumerate()
        .filter_map(|(index, name)| (index != selection.position.get()).then_some(name.as_deref()))
        .collect::<Vec<_>>();
    let actual_names = after
        .slides()
        .iter()
        .map(|slide| slide.name())
        .collect::<Vec<_>>();
    if after.slides().len() + 1 != before.slides().len() || actual_names != expected_names {
        return Err(Error::Verification);
    }
    verify_locality(source, &candidate, selection, &rewritten, previews.names())?;
    Ok(candidate)
}

fn component_archive<'a>(package: &'a Package, name: &str) -> Result<&'a Archive, Error> {
    package
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == name)
        .map(|component| component.archive())
        .ok_or(Error::InvalidSource)
}

fn verify_expected_archive_delta(
    source: &Archive,
    expected: &Archive,
    component_name: &str,
    selection: &Selection<'_>,
    limits: litchi_iwa_core::Limits,
    wire_limits: WireLimits,
) -> Result<(), Error> {
    let mut expected_index = 0usize;
    for source_object in &source.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(Error::InvalidSource)?;
        if (component_name == selection.node_component && identifier == selection.node_identifier)
            || (component_name == selection.slide_component
                && identifier == selection.slide_identifier)
        {
            continue;
        }
        let candidate = expected
            .objects
            .get(expected_index)
            .ok_or(Error::Verification)?;
        expected_index += 1;
        if identifier == selection.show_identifier && component_name == selection.show_component {
            let mut authorized = source_object.clone();
            let message = authorized
                .messages
                .get(selection.show_message_index)
                .ok_or(Error::InvalidSource)?;
            let payload = remove_show_node(&message.data, selection.node_identifier, wire_limits)?;
            authorized
                .replace_message_pruning_object_references_preserving_header_with_limits(
                    selection.show_message_index,
                    RawMessage {
                        type_: SHOW_MESSAGE_TYPE,
                        data: payload,
                    },
                    &[selection.node_identifier],
                    limits,
                )
                .map_err(map_core_error)?;
            if candidate.archive_info != authorized.archive_info
                || candidate.messages != authorized.messages
            {
                return Err(Error::Verification);
            }
        } else if identifier == selection.metadata.object_identifier
            && component_name == selection.metadata.component
        {
            let mut authorized = source_object.clone();
            authorized
                .replace_message_preserving_header_with_limits(
                    selection.metadata.message_index,
                    RawMessage {
                        type_: PACKAGE_METADATA_MESSAGE_TYPE,
                        data: selection.metadata.rewritten_payload.clone(),
                    },
                    limits,
                )
                .map_err(map_core_error)?;
            if candidate.archive_info != authorized.archive_info
                || candidate.messages != authorized.messages
            {
                return Err(Error::Verification);
            }
        } else if candidate.archive_info != source_object.archive_info
            || candidate.messages != source_object.messages
        {
            return Err(Error::Verification);
        }
    }
    if expected_index != expected.objects.len() {
        return Err(Error::Verification);
    }
    Ok(())
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: Selection<'_>,
    rewritten: &[(&str, Vec<u8>)],
    deleted_previews: &[&str],
) -> Result<(), Error> {
    let source_catalog = source.state.source.package();
    let candidate_catalog = candidate.state.source.package();
    let expected_count = source_catalog
        .len()
        .checked_sub(deleted_previews.len())
        .ok_or(Error::Verification)?;
    if candidate_catalog.len() != expected_count {
        return Err(Error::Verification);
    }
    let mut candidate_entries = candidate_catalog.iter();
    for before in source_catalog.iter() {
        if deleted_previews.contains(&before.name()) {
            continue;
        }
        let after = candidate_entries.next().ok_or(Error::Verification)?;
        if before.name() != after.name()
            || before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
        {
            return Err(Error::Verification);
        }
        if let Some((_, computed)) = rewritten.iter().find(|(name, _)| *name == before.name()) {
            if after.data() != computed
                || !super::show_settings::selected_package_member_preserved(before, after)
            {
                return Err(Error::Verification);
            }
        } else if before.data() != after.data()
            || before.metadata() != after.metadata()
            || before.raw_record().local_record() != after.raw_record().local_record()
            || before.raw_record().compressed_data() != after.raw_record().compressed_data()
            || !super::rendering_invalidation::central_record_preserved(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
            )
        {
            return Err(Error::Verification);
        }
    }
    if candidate_entries.next().is_some() {
        return Err(Error::Verification);
    }
    let before_components = source.state.source.components();
    let after_components = candidate.state.source.components();
    if before_components.len() != after_components.len() {
        return Err(Error::Verification);
    }
    for (before, after) in before_components.iter().zip(after_components.iter()) {
        if before.name() != after.name() {
            return Err(Error::Verification);
        }
        if !rewritten.iter().any(|(name, _)| *name == before.name())
            && before.archive() != after.archive()
        {
            return Err(Error::Verification);
        }
    }
    if candidate.object(selection.node_identifier).is_some()
        || candidate.object(selection.slide_identifier).is_some()
    {
        return Err(Error::Verification);
    }
    let source_order = source.private_slide_node_order().map_err(map_order_error)?;
    let target_order = candidate
        .private_slide_node_order()
        .map_err(map_order_error)?;
    let mut expected = source_order.to_vec();
    if expected.remove(selection.position.get()) != selection.node_identifier
        || expected.as_slice() != target_order.as_ref()
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn map_order_error(_error: super::SlideOrderError) -> Error {
    Error::InvalidSource
}

fn count_changed_components(before: &Package, after: &Package) -> usize {
    let before = before.state.source.components();
    let after = after.state.source.components();
    if before.len() != after.len()
        || before
            .iter()
            .zip(after.iter())
            .any(|(a, b)| a.name() != b.name())
    {
        return 0;
    }
    before
        .iter()
        .zip(after.iter())
        .filter(|(a, b)| a.archive() != b.archive())
        .count()
}

fn physical_source(package: &Package) -> Result<Arc<[u8]>, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) if source.source_is_exact() => Ok(source.shared_source()),
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
}

fn map_selector_error(_error: SlideSelectorError) -> Error {
    Error::AmbiguousSelector
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            path: _,
        } => Error::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::Objects,
                super::SemanticLimitKind::Slides => LimitKind::Slides,
                super::SemanticLimitKind::References => LimitKind::References,
                super::SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
            path: Path::Package,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            path: _,
        } => Error::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::Work,
            },
            observed: observed as u64,
            maximum: maximum as u64,
            path: Path::Package,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            observed, maximum, ..
        } => Error::LimitExceeded {
            kind: LimitKind::EntryBytes,
            observed,
            maximum,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            observed, maximum, ..
        } => Error::LimitExceeded {
            kind: LimitKind::EntryBytes,
            observed: observed as u64,
            maximum: maximum as u64,
            path: Path::Package,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::Work,
                _ => LimitKind::WireFields,
            },
            observed: observed as u64,
            maximum: limit as u64,
            path: Path::Package,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_metadata_error(error: RewriteError) -> Error {
    if let Some(limit) = error.resource_limit() {
        use litchi_iwa_protos::package_metadata_codec::RewriteLimit;
        let (kind, observed, maximum) = match limit {
            RewriteLimit::InputBytes { observed, maximum } => {
                (LimitKind::WireBytes, observed as u64, maximum as u64)
            },
            RewriteLimit::OutputBytes { observed, maximum } => {
                (LimitKind::OutputBytes, observed as u64, maximum as u64)
            },
            RewriteLimit::Fields { observed, maximum } => {
                (LimitKind::WireFields, observed as u64, maximum as u64)
            },
            RewriteLimit::Work { observed, maximum } => {
                (LimitKind::Work, observed as u64, maximum as u64)
            },
            RewriteLimit::Nesting { observed, maximum } => (
                LimitKind::WireNesting,
                u64::from(observed),
                u64::from(maximum),
            ),
            RewriteLimit::Components { observed, maximum } => {
                (LimitKind::Objects, observed as u64, maximum as u64)
            },
            RewriteLimit::References { observed, maximum } => {
                (LimitKind::References, observed as u64, maximum as u64)
            },
            RewriteLimit::Additions { observed, maximum } => {
                (LimitKind::References, observed as u64, maximum as u64)
            },
            _ => return Error::InvalidSource,
        };
        return Error::LimitExceeded {
            kind,
            observed,
            maximum,
            path: Path::PackageMetadata,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    Error::InvalidSource
}

#[cfg(test)]
mod perf_tests;
