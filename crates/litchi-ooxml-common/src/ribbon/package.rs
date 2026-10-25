//! OPC package graph discovery and mutation for shared Ribbon customizations.

use crate::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};

use super::limit::LimitKind;
use litchi_opc::TargetMode;
use litchi_opc::{
    BlobPart, OpcPackage, OwnedContentTypes, OwnedRelationships, PackURI, Part, XmlPart,
};
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
/// Returns [`Error::Signed`] for a change to a signed package, and an error when input violates OOXML
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
    let relationship_id = match existing.as_ref() {
        Some((_, id, _)) => id.clone(),
        None => next_relationship_id(package)?,
    };
    refuse_signed_change(package)?;
    package.add_part(Box::new(replacement));
    replace_root_relationship(package, &relationship_id, version.relationship(), &target)?;
    settle_unsigned_change(package)?;
    Ok(())
}

fn next_relationship_id(package: &OpcPackage) -> Result<String> {
    for number in 1..=MAX_GRAPH_LINKS {
        let candidate = format!("rId{number}");
        if package.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(Error::Limit {
        kind: crate::LimitKind::Ribbon(LimitKind::PackageGraphRelationships),
        maximum: MAX_GRAPH_LINKS,
        observed: package.rels().len(),
    })
}

fn replace_root_relationship(
    package: &mut OpcPackage,
    id: &str,
    relationship_type: &str,
    target: &str,
) -> Result<()> {
    let root = root_uri()?;
    let expected = package.source_relationships(&root, package.read_limits())?;
    let maximum = package.read_limits().max_relationship_xml_bytes();
    let replacement = expected
        .without_relationship(id, maximum)?
        .with_relationship(relationship_type, target, id, TargetMode::Internal, maximum)?;
    package.try_replace_relationships(&expected, &replacement, package.read_limits())?;
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
/// Returns [`Error::Signed`] for a change to a signed package, and an error
/// when input violates OOXML constraints, exceeds a configured bound, or an
/// underlying XML or package operation fails.
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
    let removed_parts = ribbon_part
        .iter()
        .chain(images.iter())
        .cloned()
        .collect::<Vec<_>>();
    let preserved_content_types = if removed_parts.is_empty() {
        None
    } else {
        let source = package.source_content_types(package.read_limits())?;
        Some(source.without_parts(
            &removed_parts,
            package.read_limits().max_content_types_bytes(),
        )?)
    };
    let root = root_uri()?;
    let expected = package.source_relationships(&root, package.read_limits())?;
    let replacement = expected.without_relationship(
        &relationship_id,
        package.read_limits().max_relationship_xml_bytes(),
    )?;
    package.try_replace_relationships(&expected, &replacement, package.read_limits())?;
    if let Some(part) = ribbon_part {
        let _ = package.remove_part(&part);
        for image in images {
            let _ = package.remove_part(&image);
        }
    }
    if let Some(replacement) = preserved_content_types {
        let current = package.source_content_types(package.read_limits())?;
        if current != replacement {
            package.try_replace_content_types(
                current.bytes(),
                &replacement,
                package.read_limits(),
            )?;
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
        ribbons
            .get(family)
            .map(|ui| Self::from_ui(package, ui))
            .transpose()
    }

    fn from_ui(package: &OpcPackage, ui: Ui<'_>) -> Result<Self> {
        Ok(Self {
            part: ui.part().clone(),
            id: ui.id().to_owned(),
            version: ui.version(),
            xml: package.get_part(ui.part())?.blob_arc(),
        })
    }
}

/// One exact part in the Ribbon dependency closure.
///
/// The payload and relationship token retain their source allocations.  That
/// lets a replay use the OPC preservation APIs rather than rebuilding an
/// otherwise unrelated XML member or relationship file.
#[derive(Clone, PartialEq, Eq)]
struct PartState {
    name: PackURI,
    content_type: String,
    bytes: Arc<Vec<u8>>,
    relationships: OwnedRelationships,
    retained_relationships: Option<(Arc<Vec<u8>>, bool)>,
}

#[derive(Clone, PartialEq, Eq)]
struct GraphState {
    families: [Option<OwnedUi>; 2],
    root: OwnedRelationships,
    retained_root: Option<(Arc<Vec<u8>>, bool)>,
    content_types: OwnedContentTypes,
    retained_content_types: Option<Arc<Vec<u8>>>,
    parts: Box<[PartState]>,
}

/// A source-checked Ribbon publication.
///
/// A patch may carry one or both fixed Ribbon relationship families. Keeping
/// the complete family change in one patch lets a staged edit publish one
/// commit even when it changes both legacy and modern slots.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    before: GraphState,
    after: GraphState,
    limits: Limits,
}

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field(
                "families",
                &self
                    .before
                    .families
                    .iter()
                    .zip(&self.after.families)
                    .filter(|(before, after)| before != after)
                    .count(),
            )
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether this patch preserves every Ribbon family exactly.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Return the exact inverse publication.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            limits: self.limits,
        }
    }

    /// Apply this patch to a Ribbon source with matching graph content and
    /// publish a new package.
    ///
    /// The receiver is unchanged. Unrelated package parts remain owned by the
    /// receiver; only the selected Ribbon families are planned and published.
    pub fn apply(&self, package: &OpcPackage) -> Result<Commit> {
        let hints = part_names(&self.before);
        let current = capture_state(package, &self.limits, &hints)?;
        if current != self.before {
            return Err(Error::PatchConflict);
        }
        if self.is_noop() {
            return Ok(Commit::new(
                package.clone(),
                self.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        refuse_signed_change(package)?;
        let mut candidate = package.clone();
        apply_graph_state(&mut candidate, &self.before, &self.after)?;
        let resulting = capture_state(&candidate, &self.limits, &part_names(&self.after))?;
        if !destination_matches(&resulting, &self.after) {
            return Err(Error::Internal(
                "Ribbon patch publication changed the selected graph",
            ));
        }
        settle_unsigned_change(&mut candidate)?;
        Ok(Commit::new(candidate, self.clone(), Diagnostics::CHANGED))
    }

    fn new(before: GraphState, after: GraphState, limits: Limits) -> Self {
        Self {
            before,
            after,
            limits,
        }
    }
}

/// Limits selected before a staged Ribbon edit starts.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    /// Resource ceilings used by Ribbon reads and publication.
    pub limits: Limits,
}

/// A staged Ribbon edit over an immutable OPC package source.
///
/// The source and untouched package parts remain shared until publication.
/// Dropping the edit discards its working package.
#[derive(Clone)]
pub struct Edit<'a> {
    source: &'a OpcPackage,
    working: Option<OpcPackage>,
    options: EditOptions,
}

impl std::fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("limits", &self.options.limits)
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    /// Start a staged edit with standard Ribbon limits.
    pub fn new(package: &'a OpcPackage) -> Result<Self> {
        Self::new_with_options(package, EditOptions::default())
    }

    /// Start a staged edit with caller-selected Ribbon limits.
    pub fn new_with_options(package: &'a OpcPackage, options: EditOptions) -> Result<Self> {
        Ok(Self {
            source: package,
            working: None,
            options: EditOptions {
                limits: options.limits.bounded_by_source(package.read_limits()),
            },
        })
    }

    fn current(&self) -> &OpcPackage {
        self.working.as_ref().unwrap_or(self.source)
    }

    /// Read one staged Ribbon family as owned, inert data.
    pub fn get(&self, family: Family) -> Result<Option<(Version, Vec<u8>)>> {
        Ok(
            OwnedUi::capture(self.current(), family, &self.options.limits)?
                .map(|ui| (ui.version, ui.xml.as_ref().clone())),
        )
    }

    /// Create or replace one staged Ribbon family.
    pub fn put(&mut self, version: Version, xml: Vec<u8>) -> Result<bool> {
        let family = version.family();
        if let Some(current) = OwnedUi::capture(self.current(), family, &self.options.limits)?
            && current.version == version
            && current.xml.as_slice() == xml.as_slice()
        {
            return Ok(false);
        }
        let mut candidate = self.current().clone();
        put_mut(&mut candidate, version, Arc::new(xml), &self.options.limits)?;
        self.working = Some(candidate);
        Ok(true)
    }

    /// Remove one staged Ribbon family.
    pub fn remove(&mut self, family: Family) -> Result<bool> {
        let limits = self.options.limits;
        if OwnedUi::capture(self.current(), family, &limits)?.is_none() {
            return Ok(false);
        }
        let mut candidate = self.current().clone();
        remove_mut(&mut candidate, family, &limits)?;
        self.working = Some(candidate);
        Ok(true)
    }

    /// Discard all staged operations.
    pub fn reset(&mut self) {
        self.working = None;
    }

    /// Publish all staged Ribbon families as one commit triad.
    pub fn commit(self) -> Result<Commit> {
        let before = capture_state(self.source, &self.options.limits, &[])?;
        let working = self.working.as_ref().unwrap_or(self.source);
        let after = capture_state(working, &self.options.limits, &part_names(&before))?;
        Patch::new(before, after, self.options.limits).apply(self.source)
    }
}

fn root_uri() -> Result<PackURI> {
    PackURI::new("/").map_err(|error| Error::Uri(format!("package root URI: {error}")))
}

fn part_names(state: &GraphState) -> Vec<PackURI> {
    state.parts.iter().map(|part| part.name.clone()).collect()
}

fn push_name(names: &mut Vec<PackURI>, name: PackURI) {
    if !names.iter().any(|existing| same_name(existing, &name)) {
        names.push(name);
    }
}

fn capture_state(package: &OpcPackage, limits: &Limits, hints: &[PackURI]) -> Result<GraphState> {
    let ribbons = load(package, limits)?;
    let families = [
        ribbons
            .legacy()
            .map(|ui| OwnedUi::from_ui(package, ui))
            .transpose()?,
        ribbons
            .modern()
            .map(|ui| OwnedUi::from_ui(package, ui))
            .transpose()?,
    ];

    let mut names = Vec::new();
    for hint in hints {
        if package.get_part(hint).is_ok() {
            push_name(&mut names, hint.clone());
        }
    }
    for ui in families.iter().flatten() {
        push_name(&mut names, ui.part.clone());
    }

    // Capture the complete opaque dependency closure of both family parts.
    // The traversal is bounded by the package's admitted graph; `load` has
    // already checked the Ribbon image edges, while get_part reports a
    // malformed missing internal target before any state is published.
    let mut index = 0usize;
    while index < names.len() {
        let name = names[index].clone();
        index += 1;
        let part = package.get_part(&name)?;
        for relationship in part.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!("invalid Ribbon dependency target: {error}"))
            })?;
            package.get_part(&target)?;
            push_name(&mut names, target);
        }
    }
    names.sort_unstable_by(compare_names);

    let source_limits = package.read_limits();
    let root_uri = root_uri()?;
    let root = package.source_relationships(&root_uri, source_limits)?;
    let retained_root = package.retained_relationships_xml(&root_uri)?;
    let content_types = package.source_content_types(source_limits)?;
    let retained_content_types = package.retained_content_types_xml()?;
    let mut parts = Vec::with_capacity(names.len());
    for name in names {
        let part = package.get_part(&name)?;
        let relationships = package.source_relationships(&name, source_limits)?;
        parts.push(PartState {
            name: name.clone(),
            content_type: part.content_type().to_owned(),
            bytes: part.blob_arc(),
            retained_relationships: package.retained_relationships_xml(&name)?,
            relationships,
        });
    }
    Ok(GraphState {
        families,
        root,
        retained_root,
        content_types,
        retained_content_types,
        parts: parts.into_boxed_slice(),
    })
}

fn find_part<'a>(parts: &'a [PartState], wanted: &PackURI) -> Option<&'a PartState> {
    parts.iter().find(|part| same_name(&part.name, wanted))
}

fn retained_destination_matches(
    actual: &Option<(Arc<Vec<u8>>, bool)>,
    expected: &Option<(Arc<Vec<u8>>, bool)>,
) -> bool {
    // A destination token that was authored by the staged edit has no source
    // provenance to retain. Replay installs it through the checked OPC token
    // API, which intentionally records the validated bytes as provenance;
    // that stronger proof is equivalent for publication. A retained source
    // token, however, must remain byte-identical.
    expected.is_none() || actual == expected
}

fn retained_content_types_destination_matches(
    actual: &Option<Arc<Vec<u8>>>,
    expected: &Option<Arc<Vec<u8>>>,
    content_types: &OwnedContentTypes,
) -> bool {
    match expected {
        Some(expected) => actual.as_ref() == Some(expected),
        None => actual
            .as_ref()
            .is_none_or(|actual| actual.as_slice() == content_types.bytes()),
    }
}

fn destination_matches(actual: &GraphState, expected: &GraphState) -> bool {
    if actual.families != expected.families
        || actual.root != expected.root
        || actual.content_types != expected.content_types
        || !retained_destination_matches(&actual.retained_root, &expected.retained_root)
        || !retained_content_types_destination_matches(
            &actual.retained_content_types,
            &expected.retained_content_types,
            &expected.content_types,
        )
        || actual.parts.len() != expected.parts.len()
    {
        return false;
    }
    actual
        .parts
        .iter()
        .zip(&expected.parts)
        .all(|(actual, expected)| {
            actual.name == expected.name
                && actual.content_type == expected.content_type
                && actual.bytes == expected.bytes
                && actual.relationships == expected.relationships
                && retained_destination_matches(
                    &actual.retained_relationships,
                    &expected.retained_relationships,
                )
        })
}

fn remove_part_state(package: &OpcPackage, before: &GraphState, after: &GraphState) -> Result<()> {
    let removed: Vec<&PackURI> = before
        .parts
        .iter()
        .filter(|part| find_part(&after.parts, &part.name).is_none())
        .map(|part| &part.name)
        .collect();
    for part in before.parts.iter() {
        if find_part(&after.parts, &part.name).is_some() {
            continue;
        }
        let root_inbound = package.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|target| same_name(&target, &part.name))
        });
        let part_inbound = package.iter_parts().any(|source| {
            if removed
                .iter()
                .any(|removed_name| same_name(removed_name, source.partname()))
            {
                return false;
            }
            source.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|target| same_name(&target, &part.name))
            })
        });
        if root_inbound || part_inbound {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "ribbon.apply",
                "the Ribbon graph has an opaque dependency on a removed part",
            ));
        }
    }
    Ok(())
}

fn apply_graph_state(
    package: &mut OpcPackage,
    before: &GraphState,
    after: &GraphState,
) -> Result<()> {
    let source_limits = package.read_limits();

    // Add and replace payloads through the source-preserving OPC APIs.  A
    // Ribbon operation never changes an existing part's content type; a
    // mismatch is a stale or malformed patch, not an implicit codec rewrite.
    for target in after.parts.iter() {
        match find_part(&before.parts, &target.name) {
            None => {
                if is_xml_part_content_type(&target.content_type) {
                    package.try_add_authored_xml_part_bytes(
                        target.name.clone(),
                        target.content_type.clone(),
                        Arc::clone(&target.bytes),
                        source_limits,
                    )?;
                } else {
                    // Relationship closures may contain opaque binary assets
                    // (for example Ribbon images). Restore those through the
                    // binary part API; the relationship loop below then
                    // installs any retained source `.rels` token.
                    package.try_add_part(Box::new(BlobPart::new_shared(
                        target.name.clone(),
                        target.content_type.clone(),
                        Arc::clone(&target.bytes),
                    )))?;
                }
            },
            Some(source) => {
                if source.content_type != target.content_type {
                    return Err(Error::PatchConflict);
                }
                if source.bytes.as_slice() != target.bytes.as_slice() {
                    package.try_replace_authored_xml_part_bytes(
                        &target.name,
                        source.bytes.as_slice(),
                        Arc::clone(&target.bytes),
                        source_limits,
                    )?;
                }
            },
        }
    }

    // Restore each relationship member exactly, including member presence and
    // retained comments/ordering.  This is deliberately done before removing
    // parts so newly-added targets are available to the final graph audit.
    for target in after.parts.iter() {
        let current = package.source_relationships(&target.name, source_limits)?;
        let retained = package.retained_relationships_xml(&target.name)?;
        if current != target.relationships || retained != target.retained_relationships {
            package.try_replace_relationships(&current, &target.relationships, source_limits)?;
        }
    }
    let current_root = package.source_relationships(&root_uri()?, source_limits)?;
    let retained_root = package.retained_relationships_xml(&root_uri()?)?;
    if current_root != after.root || retained_root != after.retained_root {
        package.try_replace_relationships(&current_root, &after.root, source_limits)?;
    }

    // The root and retained closure owners now describe the destination graph,
    // so only an inbound edge outside the captured Ribbon closure can keep an
    // otherwise removed part alive.
    remove_part_state(package, before, after)?;
    for source in before.parts.iter() {
        if find_part(&after.parts, &source.name).is_none() {
            let _ = package.remove_part(&source.name);
        }
    }

    let current_types = package.source_content_types(source_limits)?;
    let retained_types = package.retained_content_types_xml()?;
    if current_types != after.content_types || retained_types != after.retained_content_types {
        package.try_replace_content_types(
            current_types.bytes(),
            &after.content_types,
            source_limits,
        )?;
    }
    Ok(())
}

fn is_xml_part_content_type(content_type: &str) -> bool {
    let bytes = content_type.as_bytes();
    content_type.eq_ignore_ascii_case("application/xml")
        || (bytes.len() >= 4 && bytes[bytes.len() - 4..].eq_ignore_ascii_case(b"+xml"))
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
    Err(Error::Internal("Ribbon part-name allocation exhausted"))
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
