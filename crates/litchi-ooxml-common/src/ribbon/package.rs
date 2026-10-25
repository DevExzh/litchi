//! OPC package graph discovery and mutation for shared Ribbon customizations.

use crate::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

use super::limit::LimitKind;
use litchi_opc::{OpcPackage, PackURI, Part, XmlPart};
use std::cmp::Ordering;
use std::collections::VecDeque;
use std::sync::Arc;

use super::codec::{require_content_type, require_internal_target, validate_images, validate_xml};
use super::model::{Family, Limits, Set, Ui, Version};

pub(super) const CONTENT_TYPE: &str = "application/xml";
const PART_NAME_ATTEMPTS: usize = 10_000;
const MAX_GRAPH_LINKS: usize = 1_000_000;
const MAX_PART_NAMES: usize = 100_000;
const MAX_PART_NAME_BYTES: usize = 64 * 1024 * 1024;
const MAX_IMAGE_GC_EDGES: usize = 262_144;

/// Read both Ribbon family slots under `limits` ([`Limits::standard`] for
/// the safe defaults).
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn load<'a>(package: &'a OpcPackage, limits: &Limits) -> Result<Set<'a>> {
    let mut scanned = 0usize;
    reject_part_sourced_ribbons(package, &mut scanned)?;

    let mut result = Set::default();
    let mut images = 0usize;
    for relationship in package.rels().iter() {
        bump_graph_links(&mut scanned)?;
        let Some(family) = Family::from_relationship(relationship.reltype()) else {
            continue;
        };
        result.require_empty(family)?;
        require_internal_target(relationship, "Ribbon")?;
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid Ribbon relationship target: {error}"))
        })?;
        let part = package.get_part(&target).map_err(|error| {
            Error::Missing(format!(
                "Ribbon part '{}' does not exist: {error}",
                target.as_str()
            ))
        })?;
        require_content_type(part, CONTENT_TYPE)?;
        let version = validate_xml(part.blob(), family, limits)?;
        validate_images(package, part, &mut images, limits)?;
        result.insert(Ui::new(
            part.partname(),
            relationship.r_id(),
            version,
            part.blob(),
        ));
    }
    Ok(result)
}

/// Create or replace one Ribbon family under `limits` ([`Limits::standard`]
/// for the safe defaults).
///
/// The XML allocation is moved into the OPC part. A byte-identical update is
/// a true no-op and preserves package signatures; any other change to a
/// signed package is refused before it mutates anything (ADR 0006): remove
/// the signature explicitly first, with [`OpcPackage::unsign`] or the format
/// facade's `unsign`. Input must be UTF-8 XML.
/// # Errors
///
/// Returns [`Error::Opc`] with
/// [`SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// for a change to a signed package, and an error when input violates OOXML
/// constraints, exceeds a configured bound, or an underlying XML or package
/// operation fails.
fn put_mut(
    package: &mut OpcPackage,
    version: Version,
    xml: Arc<Vec<u8>>,
    limits: &Limits,
) -> Result<()> {
    let parsed = validate_xml(xml.as_slice(), version.family(), limits)?;
    if parsed != version {
        return Err(Error::Invalid(format!(
            "Ribbon root namespace does not match requested {version:?} version"
        )));
    }

    let existing = {
        let ribbons = load(package, limits)?;
        match ribbons.get(version.family()) {
            Some(current) if current.version() == version && current.xml() == xml.as_slice() => {
                return Ok(());
            },
            Some(current) => {
                let part = current.part().clone();
                let id = current.id().to_owned();
                let shared = has_other_inbound(package, &part, Some(&id))?;
                Some((part, id, shared))
            },
            None => None,
        }
    };

    if let Some((part, _, false)) = existing.as_ref() {
        refuse_signed_change(package)?;
        package.get_part_mut(part)?.set_blob_shared(xml);
        settle_unsigned_change(package)?;
        return Ok(());
    }

    let part = available_part_name(package, version)?;
    let target = part.as_str().trim_start_matches('/').to_owned();
    let mut replacement = XmlPart::new_shared(part, CONTENT_TYPE.to_owned(), xml);
    if let Some((shared, _, true)) = existing.as_ref() {
        let base = replacement.partname().base_uri().to_owned();
        for relationship in package.get_part(shared)?.rels().iter() {
            let image = relationship.target_partname()?;
            replacement.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                image.relative_ref(&base),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )?;
        }
    }
    let relationship_id = existing.map(|(_, id, _)| id);
    refuse_signed_change(package)?;
    if let Some(id) = relationship_id.as_deref()
        && package.rels_mut().remove(id).is_none()
    {
        return Err(Error::Relationship(format!(
            "Ribbon relationship '{id}' disappeared before commit"
        )));
    }
    package.add_part(Box::new(replacement));
    if let Some(id) = relationship_id {
        package
            .rels_mut()
            .add_relationship(version.relationship().to_owned(), target, id, false);
    } else {
        package.relate_to(&target, version.relationship());
    }
    settle_unsigned_change(package)?;
    Ok(())
}

/// Remove one Ribbon family and collect only parts that become unreferenced.
///
/// The complete deletion plan is resolved before mutation. Shared Ribbon or
/// image parts remain in the package. An absent family returns `Ok(false)` and
/// preserves signatures; removing a present family from a signed package is
/// refused before it mutates anything (ADR 0006): remove the signature
/// explicitly first, with [`OpcPackage::unsign`] or the format facade's
/// `unsign`.
/// # Errors
///
/// Returns [`Error::Opc`] with
/// [`SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// for a change to a signed package, and an error when input violates OOXML
/// constraints, exceeds a configured bound, or an underlying XML or package
/// operation fails.
fn remove_mut(package: &mut OpcPackage, family: Family, limits: &Limits) -> Result<bool> {
    let (relationship_id, ribbon_part, images) = {
        let ribbons = load(package, limits)?;
        let Some(selected) = ribbons.get(family) else {
            return Ok(false);
        };
        let ribbon_part = selected.part().clone();
        let relationship_id = selected.id().to_owned();
        let remove_part =
            !has_other_inbound(package, &ribbon_part, Some(relationship_id.as_str()))?;
        let images = if remove_part {
            let candidates = ribbon_images(package, &ribbon_part)?;
            removable_images(package, &ribbon_part, &candidates)?
        } else {
            Vec::new()
        };
        (relationship_id, remove_part.then_some(ribbon_part), images)
    };

    refuse_signed_change(package)?;
    if package.rels_mut().remove(&relationship_id).is_none() {
        return Err(Error::Relationship(format!(
            "Ribbon relationship '{relationship_id}' disappeared before commit"
        )));
    }
    if let Some(part) = ribbon_part {
        let _ = package.remove_part(&part);
        for image in images {
            let _ = package.remove_part(&image);
        }
    }
    settle_unsigned_change(package)?;
    Ok(true)
}

#[derive(Clone, PartialEq, Eq)]
struct OwnedUi {
    part: PackURI,
    id: String,
    version: Version,
    xml: Arc<Vec<u8>>,
}

impl OwnedUi {
    fn capture(package: &OpcPackage, family: Family, limits: &Limits) -> Result<Option<Self>> {
        let ribbons = load(package, limits)?;
        let Some(ui) = ribbons.get(family) else {
            return Ok(None);
        };
        Ok(Some(Self {
            part: ui.part().clone(),
            id: ui.id().to_owned(),
            version: ui.version(),
            xml: package.get_part(ui.part())?.blob_arc(),
        }))
    }
}

#[derive(Clone, PartialEq, Eq)]
struct FamilyChange {
    family: Family,
    before: Option<OwnedUi>,
    after: Option<OwnedUi>,
}

/// A source-checked Ribbon publication.
///
/// A patch may carry one or both fixed Ribbon relationship families. Keeping
/// the complete family change in one patch lets a staged edit publish one
/// commit even when it changes both legacy and modern slots.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    changes: Box<[FamilyChange]>,
    limits: Limits,
}

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field(
                "families",
                &self
                    .changes
                    .iter()
                    .filter(|change| change.before != change.after)
                    .count(),
            )
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether this patch preserves every Ribbon family exactly.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.changes
            .iter()
            .all(|change| same_content(change.before.as_ref(), change.after.as_ref()))
    }

    /// Return the exact inverse publication.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            changes: self
                .changes
                .iter()
                .rev()
                .map(|change| FamilyChange {
                    family: change.family,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            limits: self.limits,
        }
    }

    /// Apply this patch to a Ribbon source with matching family content and
    /// publish a new package.
    ///
    /// The receiver is unchanged. Unrelated package parts remain owned by the
    /// receiver; only the selected Ribbon families are planned and published.
    pub fn apply(&self, package: &OpcPackage) -> Result<Commit> {
        for change in &self.changes {
            let current = OwnedUi::capture(package, change.family, &self.limits)?;
            if !same_content(current.as_ref(), change.before.as_ref()) {
                return Err(Error::Relationship("Ribbon patch source changed".into()));
            }
        }
        if self.is_noop() {
            return Ok(Commit::new(
                package.clone(),
                self.clone(),
                litchi_core::edit::Diagnostics::UNCHANGED,
            ));
        }
        let mut candidate = package.clone();
        for change in &self.changes {
            match &change.after {
                Some(after) => put_mut(
                    &mut candidate,
                    after.version,
                    Arc::clone(&after.xml),
                    &self.limits,
                )?,
                None => {
                    remove_mut(&mut candidate, change.family, &self.limits)?;
                },
            }
        }
        for change in &self.changes {
            let resulting = OwnedUi::capture(&candidate, change.family, &self.limits)?;
            if !same_content(resulting.as_ref(), change.after.as_ref()) {
                return Err(Error::Relationship(
                    "Ribbon patch publication changed the selected graph".into(),
                ));
            }
        }
        Ok(Commit::new(
            candidate,
            self.clone(),
            litchi_core::edit::Diagnostics::CHANGED,
        ))
    }

    fn single(
        family: Family,
        before: Option<OwnedUi>,
        after: Option<OwnedUi>,
        limits: Limits,
    ) -> Self {
        Self {
            changes: vec![FamilyChange {
                family,
                before,
                after,
            }]
            .into_boxed_slice(),
            limits,
        }
    }

    fn from_families(
        before: [Option<OwnedUi>; 2],
        after: [Option<OwnedUi>; 2],
        limits: Limits,
    ) -> Self {
        let changes = [Family::Legacy, Family::Modern]
            .into_iter()
            .zip(before.into_iter().zip(after))
            .map(|(family, (before, after))| FamilyChange {
                family,
                before,
                after,
            })
            .collect();
        Self { changes, limits }
    }
}

fn same_content(left: Option<&OwnedUi>, right: Option<&OwnedUi>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => {
            left.version == right.version && left.xml.as_slice() == right.xml.as_slice()
        },
        _ => false,
    }
}

/// Publish one Ribbon family without mutating the source package.
pub fn put(
    package: &OpcPackage,
    version: Version,
    xml: Vec<u8>,
    limits: &Limits,
) -> Result<Commit> {
    let family = version.family();
    let before = OwnedUi::capture(package, family, limits)?;
    let mut candidate = package.clone();
    put_mut(&mut candidate, version, Arc::new(xml), limits)?;
    let after = OwnedUi::capture(&candidate, family, limits)?;
    let patch = Patch::single(family, before, after, *limits);
    let changed = !patch.is_noop();
    Ok(Commit::new(
        candidate,
        patch,
        litchi_core::edit::Diagnostics::new(changed),
    ))
}

/// Remove one Ribbon family without mutating the source package.
pub fn remove(package: &OpcPackage, family: Family) -> Result<Commit> {
    let limits = Limits::standard();
    let before = OwnedUi::capture(package, family, &limits)?;
    let mut candidate = package.clone();
    remove_mut(&mut candidate, family, &limits)?;
    let after = OwnedUi::capture(&candidate, family, &limits)?;
    let patch = Patch::single(family, before, after, limits);
    let changed = !patch.is_noop();
    Ok(Commit::new(
        candidate,
        patch,
        litchi_core::edit::Diagnostics::new(changed),
    ))
}

/// Limits selected before a staged Ribbon edit starts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EditOptions {
    /// Resource ceilings used by Ribbon reads and publication.
    pub limits: Limits,
}

impl Default for EditOptions {
    fn default() -> Self {
        Self {
            limits: Limits::standard(),
        }
    }
}

/// A staged Ribbon edit over an immutable OPC package source.
///
/// The source and untouched package parts remain shared until publication.
/// Dropping the edit discards its working package.
#[derive(Clone)]
pub struct Edit {
    source: OpcPackage,
    working: OpcPackage,
    options: EditOptions,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("limits", &self.options.limits)
            .finish_non_exhaustive()
    }
}

impl Edit {
    /// Start a staged edit with standard Ribbon limits.
    pub fn new(package: &OpcPackage) -> Result<Self> {
        Self::new_with_options(package, EditOptions::default())
    }

    /// Start a staged edit with caller-selected Ribbon limits.
    pub fn new_with_options(package: &OpcPackage, options: EditOptions) -> Result<Self> {
        Ok(Self {
            source: package.clone(),
            working: package.clone(),
            options,
        })
    }

    /// Read one staged Ribbon family as owned, inert data.
    pub fn get(&self, family: Family) -> Result<Option<(Version, Vec<u8>)>> {
        Ok(
            OwnedUi::capture(&self.working, family, &self.options.limits)?
                .map(|ui| (ui.version, ui.xml.as_ref().clone())),
        )
    }

    /// Create or replace one staged Ribbon family.
    pub fn put(&mut self, version: Version, xml: Vec<u8>) -> Result<bool> {
        let commit = put(&self.working, version, xml, &self.options.limits)?;
        let changed = commit.diagnostics().changed();
        self.working = commit.into_package();
        Ok(changed)
    }

    /// Remove one staged Ribbon family.
    pub fn remove(&mut self, family: Family) -> Result<bool> {
        let limits = self.options.limits;
        let before = OwnedUi::capture(&self.working, family, &limits)?;
        let mut candidate = self.working.clone();
        remove_mut(&mut candidate, family, &limits)?;
        let after = OwnedUi::capture(&candidate, family, &limits)?;
        let patch = Patch::single(family, before, after, limits);
        let changed = !patch.is_noop();
        self.working = candidate;
        Ok(changed)
    }

    /// Discard all staged operations.
    pub fn reset(&mut self) {
        self.working = self.source.clone();
    }

    /// Publish all staged Ribbon families as one commit triad.
    pub fn commit(self) -> Result<Commit> {
        let before = capture_families(&self.source, &self.options.limits)?;
        let after = capture_families(&self.working, &self.options.limits)?;
        let patch = Patch::from_families(before, after, self.options.limits);
        let changed = !patch.is_noop();
        Ok(Commit::new(
            self.working,
            patch,
            litchi_core::edit::Diagnostics::new(changed),
        ))
    }
}

fn capture_families(package: &OpcPackage, limits: &Limits) -> Result<[Option<OwnedUi>; 2]> {
    Ok([
        OwnedUi::capture(package, Family::Legacy, limits)?,
        OwnedUi::capture(package, Family::Modern, limits)?,
    ])
}

/// A published Ribbon change.
pub type Commit = litchi_core::edit::Commit<OpcPackage, Patch>;

/// Content-free diagnostics of a Ribbon commit.
pub use litchi_core::edit::Diagnostics;

fn reject_part_sourced_ribbons(package: &OpcPackage, scanned: &mut usize) -> Result<()> {
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            bump_graph_links(scanned)?;
            if Family::from_relationship(relationship.reltype()).is_some() {
                return Err(Error::Relationship(format!(
                    "Ribbon relationship '{}' must be sourced by the package, not part '{}'",
                    relationship.r_id(),
                    part.partname().as_str()
                )));
            }
        }
    }
    Ok(())
}

fn available_part_name(package: &OpcPackage, version: Version) -> Result<PackURI> {
    let existing = sorted_part_names(package)?;
    for suffix in 0..PART_NAME_ATTEMPTS {
        let path = if suffix == 0 {
            version.default_part().to_owned()
        } else {
            format!("/customUI/customUI{suffix}.xml")
        };
        let candidate = PackURI::new(&path)
            .map_err(|error| Error::Uri(format!("Ribbon part URI '{path}': {error}")))?;
        let folded = path.to_ascii_lowercase();
        if !part_name_conflicts(&existing, &folded) {
            package.validate_new_part_name(&candidate)?;
            return Ok(candidate);
        }
    }
    Err(Error::Invalid(
        "unable to allocate a unique Ribbon part name".into(),
    ))
}

fn sorted_part_names(package: &OpcPackage) -> Result<Vec<String>> {
    let mut names = Vec::with_capacity(package.part_count().min(MAX_PART_NAMES));
    let mut bytes = 0usize;
    for part in package.iter_parts() {
        let count = names.len().checked_add(1).ok_or(Error::Limit {
            kind: crate::LimitKind::Ribbon(LimitKind::PartNameAllocationScan),
            maximum: MAX_PART_NAMES,
            observed: usize::MAX,
        })?;
        if count > MAX_PART_NAMES {
            return Err(Error::Limit {
                kind: crate::LimitKind::Ribbon(LimitKind::PartNameAllocationScan),
                maximum: MAX_PART_NAMES,
                observed: count,
            });
        }
        bytes = bytes
            .checked_add(part.partname().as_str().len())
            .ok_or(Error::Limit {
                kind: crate::LimitKind::Ribbon(LimitKind::PartNameAllocationBytes),
                maximum: MAX_PART_NAME_BYTES,
                observed: usize::MAX,
            })?;
        if bytes > MAX_PART_NAME_BYTES {
            return Err(Error::Limit {
                kind: crate::LimitKind::Ribbon(LimitKind::PartNameAllocationBytes),
                maximum: MAX_PART_NAME_BYTES,
                observed: bytes,
            });
        }
        names.push(part.partname().as_str().to_ascii_lowercase());
    }
    names.sort_unstable();
    names.dedup();
    Ok(names)
}

fn part_name_conflicts(existing: &[String], candidate: &str) -> bool {
    if sorted_name_exists(existing, candidate) {
        return true;
    }
    for (index, _) in candidate.match_indices('/').skip(1) {
        if sorted_name_exists(existing, &candidate[..index]) {
            return true;
        }
    }
    let descendant_prefix = format!("{candidate}/");
    let position = existing.partition_point(|name| name.as_str() < descendant_prefix.as_str());
    existing
        .get(position)
        .is_some_and(|name| name.starts_with(&descendant_prefix))
}

fn sorted_name_exists(existing: &[String], wanted: &str) -> bool {
    existing
        .binary_search_by(|name| name.as_str().cmp(wanted))
        .is_ok()
}

fn ribbon_images(package: &OpcPackage, ribbon: &PackURI) -> Result<Vec<PackURI>> {
    let part = package.get_part(ribbon)?;
    let mut images = Vec::new();
    for relationship in part.rels().iter() {
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid Ribbon image target: {error}"))
        })?;
        let target = package.get_part(&target)?.partname();
        images.push(target.clone());
    }
    images.sort_unstable_by(compare_names);
    images.dedup_by(|left, right| same_name(left, right));
    Ok(images)
}

fn has_other_inbound(
    package: &OpcPackage,
    target: &PackURI,
    skipped_package_relationship: Option<&str>,
) -> Result<bool> {
    let mut scanned = 0usize;
    for relationship in package.rels().iter() {
        bump_graph_links(&mut scanned)?;
        if skipped_package_relationship == Some(relationship.r_id()) || relationship.is_external() {
            continue;
        }
        let related = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid package relationship target: {error}"))
        })?;
        if same_name(&related, target) {
            return Ok(true);
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            bump_graph_links(&mut scanned)?;
            if relationship.is_external() {
                continue;
            }
            let related = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid relationship target from '{}': {error}",
                    source.partname().as_str()
                ))
            })?;
            if same_name(&related, target) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn removable_images(
    package: &OpcPackage,
    removed_ribbon: &PackURI,
    candidates: &[PackURI],
) -> Result<Vec<PackURI>> {
    let names = NameIndex::new(candidates);
    let mut keep = vec![false; candidates.len()];
    let mut outgoing = vec![Vec::new(); candidates.len()];
    let mut scanned = 0usize;
    let mut edges = 0usize;

    for relationship in package.rels().iter() {
        bump_graph_links(&mut scanned)?;
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid package relationship target: {error}"))
        })?;
        if let Some(index) = names.get(&target) {
            keep[index] = true;
        }
    }

    for source in package.iter_parts() {
        if same_name(source.partname(), removed_ribbon) {
            continue;
        }
        let source_index = names.get(source.partname());
        for relationship in source.rels().iter() {
            bump_graph_links(&mut scanned)?;
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid relationship target from '{}': {error}",
                    source.partname().as_str()
                ))
            })?;
            let Some(target_index) = names.get(&target) else {
                continue;
            };
            match source_index {
                Some(source_index) => {
                    edges = edges.checked_add(1).ok_or(Error::Limit {
                        kind: crate::LimitKind::Ribbon(LimitKind::ImageGarbageCollectionEdges),
                        maximum: MAX_IMAGE_GC_EDGES,
                        observed: usize::MAX,
                    })?;
                    if edges > MAX_IMAGE_GC_EDGES {
                        return Err(Error::Limit {
                            kind: crate::LimitKind::Ribbon(LimitKind::ImageGarbageCollectionEdges),
                            maximum: MAX_IMAGE_GC_EDGES,
                            observed: edges,
                        });
                    }
                    outgoing[source_index].push(target_index);
                },
                None => keep[target_index] = true,
            }
        }
    }

    for targets in &mut outgoing {
        targets.sort_unstable();
        targets.dedup();
    }

    let mut pending: VecDeque<_> = keep
        .iter()
        .enumerate()
        .filter_map(|(index, keep)| keep.then_some(index))
        .collect();
    while let Some(source) = pending.pop_front() {
        for &target in &outgoing[source] {
            if !keep[target] {
                keep[target] = true;
                pending.push_back(target);
            }
        }
    }

    Ok(candidates
        .iter()
        .zip(keep)
        .filter(|(_, keep)| !keep)
        .map(|(candidate, _)| candidate.clone())
        .collect())
}

fn bump_graph_links(scanned: &mut usize) -> Result<()> {
    *scanned = scanned.checked_add(1).ok_or(Error::Limit {
        kind: crate::LimitKind::Ribbon(LimitKind::PackageGraphRelationships),
        maximum: MAX_GRAPH_LINKS,
        observed: usize::MAX,
    })?;
    if *scanned > MAX_GRAPH_LINKS {
        return Err(Error::Limit {
            kind: crate::LimitKind::Ribbon(LimitKind::PackageGraphRelationships),
            maximum: MAX_GRAPH_LINKS,
            observed: *scanned,
        });
    }
    Ok(())
}

struct NameIndex<'a> {
    values: &'a [PackURI],
    order: Vec<usize>,
}

impl<'a> NameIndex<'a> {
    fn new(values: &'a [PackURI]) -> Self {
        let mut order: Vec<_> = (0..values.len()).collect();
        order.sort_unstable_by(|left, right| compare_names(&values[*left], &values[*right]));
        Self { values, order }
    }

    fn get(&self, wanted: &PackURI) -> Option<usize> {
        self.order
            .binary_search_by(|index| compare_names(&self.values[*index], wanted))
            .ok()
            .map(|position| self.order[position])
    }
}

fn compare_names(left: &PackURI, right: &PackURI) -> Ordering {
    left.as_str()
        .bytes()
        .map(|byte| byte.to_ascii_lowercase())
        .cmp(right.as_str().bytes().map(|byte| byte.to_ascii_lowercase()))
}

fn same_name(left: &PackURI, right: &PackURI) -> bool {
    left.is_equivalent_to(right)
}
