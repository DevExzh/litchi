//! The private object graph of a detached owner.
//!
//! Removing a sheet or a table detaches one root from the rooted document.
//! The objects that go with it are exactly those whose every referrer goes
//! with it too: a record shared with a surviving owner (a style, a data list
//! another table also stores into) stays. The census reads the archive
//! headers' aggregate reference lists, which every native object declares,
//! so it never reinterprets a payload. Afterwards no surviving object may
//! still address a removed one, except through the edges the caller detaches
//! itself.

use std::collections::{BTreeSet, HashMap};

use super::{Fault, FaultResult, TableLeaf, Working, reserve};
use crate::TableSource;
use crate::cell_format::metadata::ENTRY_NAME as METADATA_ENTRY;

/// The source's reference census: for each object, the objects it
/// references and the objects that reference it.
pub struct Census {
    outbound: HashMap<u64, Vec<u64>>,
    inbound: HashMap<u64, Vec<u64>>,
}

/// Capacity admitted before the census starts walking object headers.
///
/// The source catalog has already passed archive ingress, so its object and
/// reference ceilings are the finite upper bounds for this phase.  Admission
/// only counts the catalog's object slots and reserves the two index maps;
/// it does not inspect graph edges.  [`Census::new`] then performs the one
/// edge walk using those pre-admitted maps and fallible per-edge growth.
pub struct Admission {
    outbound: HashMap<u64, Vec<u64>>,
    inbound: HashMap<u64, Vec<u64>>,
    max_references: usize,
}

impl Census {
    /// Admit the bounded allocations needed by a source census.
    ///
    /// Counting object slots is deliberately separate from the edge walk:
    /// the maps are reserved before any reference is traversed, while the
    /// actual graph is still visited exactly once by [`Self::new`].
    pub fn admit(working: &Working<'_, impl TableLeaf>) -> FaultResult<Admission> {
        let mut object_count = 0_usize;
        for component in working.source().catalog().iter() {
            object_count = object_count
                .checked_add(component.archive().objects.len())
                .ok_or(Fault::Limit {
                    kind: super::FaultLimit::Identifiers,
                    observed: u64::MAX,
                    maximum: u64::try_from(working.source().max_objects()).unwrap_or(u64::MAX),
                })?;
        }
        // A source accepted by ingress cannot exceed this ceiling.  Keep the
        // guard here as well because the census is also used by synthetic
        // TableSource implementations in focused tests.
        if object_count > working.source().max_objects() {
            return Err(Fault::Limit {
                kind: super::FaultLimit::Identifiers,
                observed: u64::try_from(object_count).unwrap_or(u64::MAX),
                maximum: u64::try_from(working.source().max_objects()).unwrap_or(u64::MAX),
            });
        }
        let mut outbound = HashMap::new();
        outbound
            .try_reserve(object_count)
            .map_err(|_| super::allocation(object_count))?;
        let mut inbound = HashMap::new();
        // A target can occur at most once per source reference.  Reserving
        // the smaller of the object and reference ceilings avoids turning a
        // large configurable reference limit into an eager unbounded map.
        let inbound_capacity = object_count.min(working.source().max_references());
        inbound
            .try_reserve(inbound_capacity)
            .map_err(|_| super::allocation(inbound_capacity))?;
        Ok(Admission {
            outbound,
            inbound,
            max_references: working.source().max_references(),
        })
    }

    /// Build the census of every object of `working`'s source.
    ///
    /// Callers that have already admitted a phase can use
    /// [`Self::from_admission`] to make that ordering explicit.
    pub fn new(working: &Working<'_, impl TableLeaf>) -> FaultResult<Self> {
        let admission = Self::admit(working)?;
        Self::from_admission(working, admission)
    }

    /// Build the census after a caller has admitted its bounded phase.
    pub fn from_admission(
        working: &Working<'_, impl TableLeaf>,
        admission: Admission,
    ) -> FaultResult<Self> {
        let Admission {
            mut outbound,
            mut inbound,
            max_references,
        } = admission;
        let mut reference_count = 0_usize;
        for component in working.source().catalog().iter() {
            for object in &component.archive().objects {
                let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                let mut targets: Vec<u64> = Vec::new();
                for info in &object.archive_info.message_infos {
                    for target in &info.object_references {
                        if !targets.contains(target) {
                            reference_count =
                                reference_count.checked_add(1).ok_or(Fault::Limit {
                                    kind: super::FaultLimit::Identifiers,
                                    observed: u64::MAX,
                                    maximum: u64::try_from(max_references).unwrap_or(u64::MAX),
                                })?;
                            if reference_count > max_references {
                                return Err(Fault::Limit {
                                    kind: super::FaultLimit::Identifiers,
                                    observed: u64::try_from(reference_count).unwrap_or(u64::MAX),
                                    maximum: u64::try_from(max_references).unwrap_or(u64::MAX),
                                });
                            }
                            reserve(&mut targets, 1)?;
                            targets.push(*target);
                        }
                    }
                    for field in &info.field_infos {
                        for target in &field.object_references {
                            if !targets.contains(target) {
                                reference_count =
                                    reference_count.checked_add(1).ok_or(Fault::Limit {
                                        kind: super::FaultLimit::Identifiers,
                                        observed: u64::MAX,
                                        maximum: u64::try_from(max_references).unwrap_or(u64::MAX),
                                    })?;
                                if reference_count > max_references {
                                    return Err(Fault::Limit {
                                        kind: super::FaultLimit::Identifiers,
                                        observed: u64::try_from(reference_count)
                                            .unwrap_or(u64::MAX),
                                        maximum: u64::try_from(max_references).unwrap_or(u64::MAX),
                                    });
                                }
                                reserve(&mut targets, 1)?;
                                targets.push(*target);
                            }
                        }
                    }
                }
                for target in &targets {
                    if !inbound.contains_key(target) {
                        inbound.try_reserve(1).map_err(|_| super::allocation(1))?;
                    }
                    let referrers = inbound.entry(*target).or_default();
                    if !referrers.contains(&identifier) {
                        reserve(referrers, 1)?;
                        referrers.push(identifier);
                    }
                }
                outbound.insert(identifier, targets);
            }
        }
        Ok(Self { outbound, inbound })
    }

    /// The objects that reference `identifier`.
    pub fn referrers(&self, identifier: u64) -> &[u64] {
        self.inbound.get(&identifier).map_or(&[], Vec::as_slice)
    }

    /// The objects `identifier` references.
    pub fn targets(&self, identifier: u64) -> &[u64] {
        self.outbound.get(&identifier).map_or(&[], Vec::as_slice)
    }

    /// The private closure of `roots`: the roots and every object all of
    /// whose referrers are in the closure. `protected` objects (the package
    /// root, the registry) never join.
    pub fn private_closure(
        &self,
        working: &Working<'_, impl TableLeaf>,
        roots: &[u64],
    ) -> FaultResult<BTreeSet<u64>> {
        let mut protected = vec![1_u64];
        let registry = working
            .archive(METADATA_ENTRY)
            .map_or(&[][..], |archive| archive.objects.as_slice());
        for object in registry {
            protected.push(object.archive_info.identifier.ok_or(Fault::Invalid)?);
        }
        let mut closure: BTreeSet<u64> = roots.iter().copied().collect();
        loop {
            let mut added = Vec::new();
            for member in &closure {
                for target in self.targets(*member) {
                    if closure.contains(target)
                        || added.contains(target)
                        || protected.contains(target)
                        || !self.outbound.contains_key(target)
                    {
                        continue;
                    }
                    if self
                        .referrers(*target)
                        .iter()
                        .all(|referrer| closure.contains(referrer))
                    {
                        reserve(&mut added, 1)?;
                        added.push(*target);
                    }
                }
            }
            if added.is_empty() {
                return Ok(closure);
            }
            closure.extend(added);
        }
    }

    /// Refuse when an object outside `closure` still references a member,
    /// other than through one of the `detached` (referrer, target) edges.
    pub fn require_detached(
        &self,
        closure: &BTreeSet<u64>,
        detached: &[(u64, u64)],
    ) -> FaultResult<()> {
        for member in closure {
            for referrer in self.referrers(*member) {
                if closure.contains(referrer) || detached.contains(&(*referrer, *member)) {
                    continue;
                }
                return Err(Fault::Dependency);
            }
        }
        Ok(())
    }
}

/// Prove that no object of `candidate` is a removed identifier or still
/// references one.
pub fn verify_removed(candidate: &dyn TableSource, removed: &BTreeSet<u64>) -> FaultResult<()> {
    if removed.is_empty() {
        return Ok(());
    }
    for component in candidate.catalog().iter() {
        for object in &component.archive().objects {
            if object
                .archive_info
                .identifier
                .is_some_and(|identifier| removed.contains(&identifier))
            {
                return Err(Fault::Verification);
            }
            for info in &object.archive_info.message_infos {
                let stale = info
                    .object_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| field.object_references.iter()),
                    )
                    .any(|identifier| removed.contains(identifier));
                if stale {
                    return Err(Fault::Verification);
                }
            }
        }
    }
    Ok(())
}
