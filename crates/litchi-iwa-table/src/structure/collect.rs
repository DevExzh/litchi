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

impl Census {
    /// Build the census of every object of `working`'s source.
    pub fn new(working: &Working<'_, impl TableLeaf>) -> FaultResult<Self> {
        let mut outbound: HashMap<u64, Vec<u64>> = HashMap::new();
        let mut inbound: HashMap<u64, Vec<u64>> = HashMap::new();
        for component in working.source().catalog().iter() {
            for object in &component.archive().objects {
                let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                let mut targets: Vec<u64> = Vec::new();
                for info in &object.archive_info.message_infos {
                    for target in &info.object_references {
                        if !targets.contains(target) {
                            reserve(&mut targets, 1)?;
                            targets.push(*target);
                        }
                    }
                }
                for target in &targets {
                    let referrers = inbound.entry(*target).or_default();
                    if !referrers.contains(&identifier) {
                        reserve(referrers, 1)?;
                        referrers.push(identifier);
                    }
                }
                outbound
                    .try_reserve(1)
                    .map_err(|_error| super::allocation(1))?;
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
