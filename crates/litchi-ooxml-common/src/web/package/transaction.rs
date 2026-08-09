use super::super::codec::{invalid, limit};
use super::super::model::Limits;
use super::super::{
    Arc, BTreeSet, Error, HashMap, HashSet, OpcPackage, PackURI, Part, Result,
    TASK_PANES_RELATIONSHIP, VecDeque,
};
use super::{PlannedPart, PlannedRelationship, fold_part_name, folded_name_conflicts};
use litchi_opc::{BlobPart, TargetMode};
/// Opaque, exact, reversible task-pane graph transaction.
///
/// A patch records the precise source and destination state of every affected
/// relationship and part. Applying a stale patch fails before mutation. Empty
/// patches preserve package signatures; a changed package is always unsigned.
#[must_use = "a planned Web Extensions patch has no effect until it is applied"]
#[derive(Clone, Default)]
pub struct Patch {
    pub(in crate::web) root: Option<RootChange>,
    pub(in crate::web) parts: Box<[PartChange]>,
    pub(in crate::web) protection: Option<ProtectionGuard>,
}

pub(in crate::web) struct PatchPlan {
    pub(in crate::web) before: PlannedGraph,
    pub(in crate::web) after: PlannedGraph,
    pub(in crate::web) parts: Vec<PlannedPart>,
    pub(in crate::web) deletions: Vec<PackURI>,
    pub(in crate::web) limits: Limits,
}

pub(in crate::web) struct PlannedGraph {
    pub(in crate::web) root: Option<RelationshipState>,
    pub(in crate::web) owned_parts: Vec<PackURI>,
}

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("empty", &self.is_empty())
            .field("affected_parts", &self.parts.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    pub(in crate::web) fn planned(package: &OpcPackage, plan: PatchPlan) -> Result<Self> {
        let PatchPlan {
            before,
            after,
            parts: planned,
            deletions,
            limits,
        } = plan;
        let affected_parts = planned
            .len()
            .checked_add(deletions.len())
            .ok_or(Error::Limit {
                resource: "Web Extensions patch parts",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let mut parts = Vec::with_capacity(affected_parts);
        let mut planned_names = HashSet::with_capacity(planned.len());
        for part in planned {
            let name = part.name.clone();
            planned_names.insert(fold_part_name(&name));
            let before = package.get_part(&name).ok().map(PartState::capture);
            parts.push(PartChange {
                name,
                before,
                after: Some(PartState::from_planned(part)),
            });
        }
        for name in deletions {
            if planned_names.contains(&fold_part_name(&name)) {
                return invalid("planned Web Extensions part is also marked for deletion".into());
            }
            let before = package
                .get_part(&name)
                .ok()
                .map(PartState::capture)
                .ok_or_else(|| {
                    Error::Missing("Web Extensions patch source part disappeared".into())
                })?;
            parts.push(PartChange {
                name,
                before: Some(before),
                after: None,
            });
        }
        parts.sort_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        let protection = ProtectionGuard {
            source: GraphScope {
                root_relationship_id: before.root.as_ref().map(|root| root.id.clone()),
                owned_parts: before.owned_parts.into_boxed_slice(),
            },
            destination: GraphScope {
                root_relationship_id: after.root.as_ref().map(|root| root.id.clone()),
                owned_parts: after.owned_parts.into_boxed_slice(),
            },
            limits,
        };
        let patch = Self {
            root: Some(RootChange {
                before: before.root,
                after: after.root,
            }),
            parts: parts.into_boxed_slice(),
            protection: Some(protection),
        };
        if patch.has_changes() {
            Ok(patch)
        } else {
            Ok(Self::default())
        }
    }

    /// Return whether this patch makes no changes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.has_changes()
    }

    /// Build the exact inverse without copying part payloads.
    #[must_use = "the inverse has no effect until it is applied"]
    pub fn inverse(&self) -> Self {
        Self {
            root: self.root.as_ref().map(|root| RootChange {
                before: root.after.clone(),
                after: root.before.clone(),
            }),
            parts: self
                .parts
                .iter()
                .map(|part| PartChange {
                    name: part.name.clone(),
                    before: part.after.clone(),
                    after: part.before.clone(),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            protection: self.protection.as_ref().map(ProtectionGuard::inverse),
        }
    }

    /// Apply this patch after checking its exact source graph.
    ///
    /// Returns `true` when the package changed. Source checks, destination-name
    /// checks, and relationship staging all finish before the first mutation.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<bool> {
        if self.is_empty() {
            return Ok(false);
        }
        self.validate_source(package)?;
        self.validate_protection(package)?;
        self.validate_destination_names(package)?;
        let staged = self
            .parts
            .iter()
            .filter(|part| part.before != part.after)
            .filter_map(|part| {
                part.after
                    .as_ref()
                    .map(|after| after.blob_part(part.name.clone()))
            })
            .collect::<Result<Vec<_>>>()?;

        for part in staged {
            package.add_part(Box::new(part));
        }
        if let Some(root) = self.root.as_ref().filter(|root| root.before != root.after) {
            if let Some(before) = &root.before {
                package.rels_mut().remove(&before.id);
            }
            if let Some(after) = &root.after {
                package.rels_mut().add_relationship(
                    after.relationship_type.clone(),
                    after.target.clone(),
                    after.id.clone(),
                    after.external,
                );
            }
        }
        for part in self
            .parts
            .iter()
            .filter(|part| part.before != part.after && part.after.is_none())
        {
            package.remove_part(&part.name);
        }
        package.unsign();
        Ok(true)
    }

    pub(in crate::web) fn has_changes(&self) -> bool {
        self.root
            .as_ref()
            .is_some_and(|root| root.before != root.after)
            || self.parts.iter().any(|part| part.before != part.after)
    }

    pub(in crate::web) fn validate_source(&self, package: &OpcPackage) -> Result<()> {
        if let Some(root) = &self.root {
            let mut task_relationships = package
                .rels()
                .iter()
                .filter(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP);
            let first = task_relationships.next();
            let unique = task_relationships.next().is_none();
            let root_matches = match &root.before {
                Some(before) => first.is_some_and(|actual| before.matches(actual)) && unique,
                None => first.is_none(),
            };
            if !root_matches {
                return Err(Error::Relationship(
                    "Web Extensions patch source relationship changed".into(),
                ));
            }
            if let Some(after) = &root.after {
                let reuses_source_id = root
                    .before
                    .as_ref()
                    .is_some_and(|before| before.id == after.id);
                if !reuses_source_id && package.rels().get(&after.id).is_some() {
                    return Err(Error::Relationship(
                        "Web Extensions patch destination relationship is occupied".into(),
                    ));
                }
            }
        }
        for change in &self.parts {
            let actual = package.get_part(&change.name).ok();
            let matches = match (&change.before, actual) {
                (Some(before), Some(actual)) => before.matches(actual),
                (None, None) => true,
                _ => false,
            };
            if !matches {
                return Err(Error::Relationship(
                    "Web Extensions patch source part changed".into(),
                ));
            }
        }
        Ok(())
    }

    pub(in crate::web) fn validate_protection(&self, package: &OpcPackage) -> Result<()> {
        let Some(guard) = &self.protection else {
            return Ok(());
        };
        let changed_existing: HashSet<_> = self
            .parts
            .iter()
            .filter(|part| part.before.is_some() && part.before != part.after)
            .map(|part| fold_part_name(&part.name))
            .collect();
        if !changed_existing.is_empty() {
            let protected = protected_parts(package, &guard.source, &guard.limits)?;
            if changed_existing.iter().any(|name| protected.contains(name)) {
                return Err(Error::Relationship(
                    "Web Extensions patch would change a newly shared source part".into(),
                ));
            }
        }

        let changed_destinations: HashSet<_> = self
            .parts
            .iter()
            .filter(|part| part.after.is_some() && part.before != part.after)
            .map(|part| fold_part_name(&part.name))
            .collect();
        if !changed_destinations.is_empty() {
            let protected = protected_parts(package, &guard.destination, &guard.limits)?;
            if changed_destinations
                .iter()
                .any(|name| protected.contains(name))
            {
                return Err(Error::Relationship(
                    "Web Extensions patch would create a newly shared destination part".into(),
                ));
            }
        }
        Ok(())
    }

    pub(in crate::web) fn validate_destination_names(&self, package: &OpcPackage) -> Result<()> {
        let changed = self
            .parts
            .iter()
            .filter(|change| change.before != change.after && change.after.is_some());
        let mut destinations = BTreeSet::new();
        let mut replacements = HashMap::new();
        for change in changed {
            let folded = fold_part_name(&change.name);
            if folded_name_conflicts(&destinations, &folded) {
                return invalid("Web Extensions patch destination parts conflict".into());
            }
            destinations.insert(folded.clone());
            if change.before.is_some() {
                replacements.insert(folded, &change.name);
            }
        }
        if destinations.is_empty() {
            return Ok(());
        }
        for existing in package.iter_parts() {
            let folded = fold_part_name(existing.partname());
            if !folded_name_conflicts(&destinations, &folded) {
                continue;
            }
            let is_replaced_source = replacements
                .get(&folded)
                .is_some_and(|name| *name == existing.partname());
            if !is_replaced_source {
                return invalid("Web Extensions patch destination part is occupied".into());
            }
        }
        Ok(())
    }
}

#[derive(Clone, PartialEq, Eq)]
pub(in crate::web) struct RootChange {
    pub(in crate::web) before: Option<RelationshipState>,
    pub(in crate::web) after: Option<RelationshipState>,
}

#[derive(Clone)]
pub(in crate::web) struct ProtectionGuard {
    pub(in crate::web) source: GraphScope,
    pub(in crate::web) destination: GraphScope,
    pub(in crate::web) limits: Limits,
}

impl ProtectionGuard {
    pub(in crate::web) fn inverse(&self) -> Self {
        Self {
            source: self.destination.clone(),
            destination: self.source.clone(),
            limits: self.limits,
        }
    }
}

#[derive(Clone)]
pub(in crate::web) struct GraphScope {
    pub(in crate::web) root_relationship_id: Option<String>,
    pub(in crate::web) owned_parts: Box<[PackURI]>,
}

pub(in crate::web) fn protected_parts(
    package: &OpcPackage,
    scope: &GraphScope,
    limits: &Limits,
) -> Result<HashSet<String>> {
    let part_count = package.part_count();
    if part_count > limits.package_parts {
        return limit("package parts", limits.package_parts, part_count);
    }
    if scope.owned_parts.is_empty() {
        return Ok(HashSet::new());
    }

    let mut names = Vec::with_capacity(scope.owned_parts.len());
    let mut by_name = HashMap::with_capacity(scope.owned_parts.len());
    for name in &scope.owned_parts {
        let folded = fold_part_name(name);
        if by_name.insert(folded.clone(), names.len()).is_some() {
            return invalid("duplicate part in Web Extensions patch protection scope".into());
        }
        names.push(folded);
    }

    let mut outbound = vec![Vec::new(); names.len()];
    let mut protected = vec![false; names.len()];
    let mut queue = VecDeque::new();
    let mut relationships = 0usize;

    for relationship in package.rels().iter() {
        charge_patch_relationship(&mut relationships, limits)?;
        if relationship.is_external() {
            continue;
        }
        let Ok(target) = relationship.target_partname() else {
            continue;
        };
        let Some(&target) = by_name.get(&fold_part_name(&target)) else {
            continue;
        };
        if scope.root_relationship_id.as_deref() != Some(relationship.r_id()) && !protected[target]
        {
            protected[target] = true;
            queue.push_back(target);
        }
    }

    for part in package.iter_parts() {
        let source = by_name.get(&fold_part_name(part.partname())).copied();
        for relationship in part.rels().iter() {
            charge_patch_relationship(&mut relationships, limits)?;
            if relationship.is_external() {
                continue;
            }
            let Ok(target_name) = relationship.target_partname() else {
                continue;
            };
            let Some(&target) = by_name.get(&fold_part_name(&target_name)) else {
                continue;
            };
            if let Some(source) = source {
                outbound[source].push(target);
            } else if !protected[target] {
                protected[target] = true;
                queue.push_back(target);
            }
        }
    }

    while let Some(source) = queue.pop_front() {
        for &target in &outbound[source] {
            if !protected[target] {
                protected[target] = true;
                queue.push_back(target);
            }
        }
    }

    Ok(names
        .into_iter()
        .enumerate()
        .filter_map(|(index, name)| protected[index].then_some(name))
        .collect())
}

pub(in crate::web) fn charge_patch_relationship(
    relationships: &mut usize,
    limits: &Limits,
) -> Result<()> {
    *relationships = relationships.checked_add(1).ok_or(Error::Limit {
        resource: "package relationships",
        max: limits.package_relationships,
        actual: usize::MAX,
    })?;
    if *relationships > limits.package_relationships {
        return limit(
            "package relationships",
            limits.package_relationships,
            *relationships,
        );
    }
    Ok(())
}

#[derive(Clone, PartialEq, Eq)]
pub(in crate::web) struct PartChange {
    pub(in crate::web) name: PackURI,
    pub(in crate::web) before: Option<PartState>,
    pub(in crate::web) after: Option<PartState>,
}

#[derive(Clone, PartialEq, Eq)]
pub(in crate::web) struct PartState {
    pub(in crate::web) content_type: String,
    pub(in crate::web) data: Arc<Vec<u8>>,
    pub(in crate::web) relationships: Box<[RelationshipState]>,
}

impl PartState {
    pub(in crate::web) fn capture(part: &dyn Part) -> Self {
        let mut relationships = part
            .rels()
            .iter()
            .map(RelationshipState::capture)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        Self {
            content_type: part.content_type().to_owned(),
            data: part.blob_arc(),
            relationships: relationships.into_boxed_slice(),
        }
    }

    pub(in crate::web) fn from_planned(part: PlannedPart) -> Self {
        let mut relationships = part
            .relationships
            .into_iter()
            .map(RelationshipState::from_planned)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        Self {
            content_type: part.content_type,
            data: part.data,
            relationships: relationships.into_boxed_slice(),
        }
    }

    pub(in crate::web) fn matches(&self, part: &dyn Part) -> bool {
        self.content_type == part.content_type()
            && self.data.as_slice() == part.blob()
            && self.relationships.len() == part.rels().len()
            && self.relationships.iter().all(|expected| {
                part.rels()
                    .get(&expected.id)
                    .is_some_and(|actual| expected.matches(actual))
            })
    }

    pub(in crate::web) fn blob_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.data));
        for relationship in &self.relationships {
            part.rels_mut().try_add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                if relationship.external {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )?;
        }
        Ok(part)
    }
}

#[derive(Clone, PartialEq, Eq)]
pub(in crate::web) struct RelationshipState {
    pub(in crate::web) id: String,
    pub(in crate::web) relationship_type: String,
    pub(in crate::web) target: String,
    pub(in crate::web) external: bool,
}

impl RelationshipState {
    pub(in crate::web) fn capture(relationship: &litchi_opc::Relationship) -> Self {
        Self {
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        }
    }

    pub(in crate::web) fn from_planned(relationship: PlannedRelationship) -> Self {
        Self {
            id: relationship.id,
            relationship_type: relationship.relationship_type,
            target: relationship.target,
            external: relationship.external,
        }
    }

    pub(in crate::web) fn matches(&self, relationship: &litchi_opc::Relationship) -> bool {
        self.id == relationship.r_id()
            && self.relationship_type == relationship.reltype()
            && self.target == relationship.target_ref()
            && self.external == relationship.is_external()
    }
}
