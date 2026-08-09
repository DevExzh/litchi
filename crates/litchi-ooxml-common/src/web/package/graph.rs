use super::super::codec::{invalid, limit};
use super::super::model::{Limits, OperationBudget};
use super::super::{Arc, BTreeSet, Error, HashMap, HashSet, OpcPackage, PackURI, Result, VecDeque};
use super::fold_part_name;
#[derive(Debug)]
pub(in crate::web) struct ExistingAddInGraph {
    pub(in crate::web) root_relationship_id: String,
    pub(in crate::web) task_panes_name: PackURI,
    pub(in crate::web) extensions_by_relationship: HashMap<String, PackURI>,
    pub(in crate::web) owned_parts: Vec<PackURI>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::web) struct PlannedRelationship {
    pub(in crate::web) id: String,
    pub(in crate::web) relationship_type: String,
    pub(in crate::web) target: String,
    pub(in crate::web) external: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::web) struct PlannedPart {
    pub(in crate::web) name: PackURI,
    pub(in crate::web) content_type: String,
    pub(in crate::web) data: Arc<Vec<u8>>,
    pub(in crate::web) relationships: Vec<PlannedRelationship>,
}

#[derive(Debug)]
pub(in crate::web) struct IndexedInbound {
    pub(in crate::web) source: Option<usize>,
    pub(in crate::web) relationship_id: String,
}

#[derive(Debug)]
pub(in crate::web) struct IndexedPart {
    pub(in crate::web) name: PackURI,
    pub(in crate::web) outbound: Vec<usize>,
    pub(in crate::web) inbound: Vec<IndexedInbound>,
}

/// One bounded, ASCII-case-folded view of package membership and internal edges.
#[derive(Debug)]
pub(in crate::web) struct PackageGraphIndex {
    pub(in crate::web) parts: Vec<IndexedPart>,
    pub(in crate::web) by_folded: HashMap<String, usize>,
    pub(in crate::web) occupied: BTreeSet<String>,
    pub(in crate::web) relationships: usize,
}

impl PackageGraphIndex {
    pub(in crate::web) fn build(
        package: &OpcPackage,
        limits: &Limits,
        budget: &mut OperationBudget,
    ) -> Result<Self> {
        let part_count = package.part_count();
        if part_count > limits.package_parts {
            return limit("package parts", limits.package_parts, part_count);
        }
        let mut parts: Vec<IndexedPart> = Vec::with_capacity(part_count);
        let mut by_folded = HashMap::with_capacity(part_count);
        let mut occupied = BTreeSet::new();
        for part in package.iter_parts() {
            let metadata_bytes = part
                .partname()
                .as_str()
                .len()
                .checked_add(part.content_type().len())
                .ok_or(Error::Limit {
                    resource: "indexed web extension package metadata bytes",
                    max: limits.total_string_bytes,
                    actual: usize::MAX,
                })?;
            budget.charge_metadata(metadata_bytes, 4, limits)?;
            let folded = fold_part_name(part.partname());
            if let Some(index) = by_folded.insert(folded.clone(), parts.len()) {
                return invalid(format!(
                    "ASCII-case-equivalent package parts '{}' and '{}' coexist",
                    parts[index].name.as_str(),
                    part.partname().as_str()
                ));
            }
            occupied.insert(folded);
            parts.push(IndexedPart {
                name: part.partname().clone(),
                outbound: Vec::new(),
                inbound: Vec::new(),
            });
        }

        let mut value = Self {
            parts,
            by_folded,
            occupied,
            relationships: 0,
        };
        for relationship in package.rels().iter() {
            value.record_relationship(None, relationship, limits, budget)?;
        }
        for part in package.iter_parts() {
            let source = value
                .index_of(part.partname())
                .ok_or_else(|| Error::Missing(part.partname().to_string()))?;
            for relationship in part.rels().iter() {
                value.record_relationship(Some(source), relationship, limits, budget)?;
            }
        }
        Ok(value)
    }

    pub(in crate::web) fn record_relationship(
        &mut self,
        source: Option<usize>,
        relationship: &litchi_opc::Relationship,
        limits: &Limits,
        budget: &mut OperationBudget,
    ) -> Result<()> {
        let metadata_bytes = relationship
            .r_id()
            .len()
            .checked_add(relationship.reltype().len())
            .and_then(|bytes| bytes.checked_add(relationship.target_ref().len()))
            .ok_or(Error::Limit {
                resource: "indexed web extension package metadata bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?;
        budget.charge_metadata(metadata_bytes, 3, limits)?;
        self.relationships = self.relationships.checked_add(1).ok_or(Error::Limit {
            resource: "package relationships",
            max: limits.package_relationships,
            actual: usize::MAX,
        })?;
        if self.relationships > limits.package_relationships {
            return limit(
                "package relationships",
                limits.package_relationships,
                self.relationships,
            );
        }
        if relationship.is_external() {
            return Ok(());
        }
        let Ok(target) = relationship.target_partname() else {
            // Web graph relationships are rejected with context by their callers.
            return Ok(());
        };
        let Some(target) = self.index_of(&target) else {
            return Ok(());
        };
        if let Some(source) = source {
            self.parts[source].outbound.push(target);
        }
        self.parts[target].inbound.push(IndexedInbound {
            source,
            relationship_id: relationship.r_id().to_owned(),
        });
        Ok(())
    }

    pub(in crate::web) fn index_of(&self, name: &PackURI) -> Option<usize> {
        self.by_folded.get(&fold_part_name(name)).copied()
    }

    pub(in crate::web) fn canonical(&self, name: &PackURI) -> Option<&PackURI> {
        self.index_of(name).map(|index| &self.parts[index].name)
    }

    pub(in crate::web) fn contains(&self, name: &PackURI) -> bool {
        self.index_of(name).is_some()
    }

    pub(in crate::web) fn conflicts(&self, candidate: &PackURI) -> bool {
        let folded = fold_part_name(candidate);
        if self.occupied.contains(&folded) {
            return true;
        }
        let mut ancestor = folded.as_str();
        while let Some(index) = ancestor.rfind('/') {
            if index == 0 {
                break;
            }
            ancestor = &ancestor[..index];
            if self.occupied.contains(ancestor) {
                return true;
            }
        }
        let descendant_prefix = format!("{folded}/");
        self.occupied
            .range(descendant_prefix.clone()..)
            .next()
            .is_some_and(|name| name.starts_with(&descendant_prefix))
    }

    pub(in crate::web) fn protected_closure(
        &self,
        owned_parts: &[PackURI],
        root_relationship_id: &str,
    ) -> HashSet<String> {
        let owned: HashSet<_> = owned_parts
            .iter()
            .filter_map(|name| self.index_of(name))
            .collect();
        let mut queue = VecDeque::new();
        let mut protected = HashSet::new();
        for &index in &owned {
            let has_external_ingress =
                self.parts[index]
                    .inbound
                    .iter()
                    .any(|inbound| match inbound.source {
                        None => inbound.relationship_id != root_relationship_id,
                        Some(source) => !owned.contains(&source),
                    });
            if has_external_ingress && protected.insert(index) {
                queue.push_back(index);
            }
        }
        while let Some(source) = queue.pop_front() {
            for &target in &self.parts[source].outbound {
                if owned.contains(&target) && protected.insert(target) {
                    queue.push_back(target);
                }
            }
        }
        protected
            .into_iter()
            .map(|index| fold_part_name(&self.parts[index].name))
            .collect()
    }
}
