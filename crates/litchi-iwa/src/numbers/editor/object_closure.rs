//! The private object closure of a detached Numbers owner.
//!
//! Host defect fixed with record 0370g: `remove_table` removed the table
//! drawable, its model, its cell storage and its formula owner, but left
//! every other object only the table addressed: summary and category
//! models with their own storage, hidden-state and sort owners, pivot
//! records and table-private styles. Numbers keeps no unreferenced object;
//! the closure below is what goes with a table, as the native app and the
//! focused crate remove it.

use super::*;
use std::collections::BTreeSet;

/// `roots` and every object all of whose referrers are in the closure,
/// read from the archive headers' aggregate reference lists. The document
/// root and the package metadata objects never join.
pub(super) fn private_closure(package: &HostPackage, roots: &[u64]) -> Result<BTreeSet<u64>> {
    let mut outbound: HashMap<u64, Vec<u64>> = HashMap::new();
    let mut inbound: HashMap<u64, Vec<u64>> = HashMap::new();
    let mut protected: HashSet<u64> = HashSet::from([1]);
    for name in package.iwa_entry_names() {
        package.with_parsed_archive(name, |archive| {
            for object in &archive.objects {
                let identifier = object.archive_info.identifier.ok_or_else(|| {
                    Error::Archive(format!("Object in {name} has no archive identifier"))
                })?;
                if name == "Index/Metadata.iwa" {
                    protected.insert(identifier);
                }
                let mut targets: Vec<u64> = Vec::new();
                for info in &object.archive_info.message_infos {
                    for target in &info.object_references {
                        if !targets.contains(target) {
                            targets.push(*target);
                        }
                    }
                }
                for target in &targets {
                    let referrers = inbound.entry(*target).or_default();
                    if !referrers.contains(&identifier) {
                        referrers.push(identifier);
                    }
                }
                outbound.insert(identifier, targets);
            }
            Ok(())
        })?;
    }
    let mut closure: BTreeSet<u64> = roots.iter().copied().collect();
    loop {
        let mut added = Vec::new();
        for member in &closure {
            for target in outbound.get(member).map_or(&[][..], Vec::as_slice) {
                if closure.contains(target)
                    || added.contains(target)
                    || protected.contains(target)
                    || !outbound.contains_key(target)
                {
                    continue;
                }
                if inbound
                    .get(target)
                    .map_or(&[][..], Vec::as_slice)
                    .iter()
                    .all(|referrer| closure.contains(referrer))
                {
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
