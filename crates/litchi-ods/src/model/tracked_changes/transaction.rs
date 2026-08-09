//! Source-bound tracked-change snapshots, inert edits, and reversible patches.

use super::model::{RelationKind, Resources};
use super::{Acceptance, Change, Changes, Limits, codec};
use litchi_core::{Error, Result};
use std::{collections::HashMap, mem, sync::Arc};

/// An immutable, presence-aware view of one exact ODS `content.xml` source.
#[derive(Clone, Debug)]
pub struct Snapshot {
    source: Arc<str>,
    limits: Limits,
    map: Arc<codec::SourceMap>,
}

impl Snapshot {
    /// Parse an exact `content.xml` source with default resource limits.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse(source: impl Into<Arc<str>>) -> Result<Self> {
        Self::parse_with_limits(source, Limits::default())
    }

    /// Parse an exact `content.xml` source with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse_with_limits(source: impl Into<Arc<str>>, limits: Limits) -> Result<Self> {
        let source = source.into();
        if source.len() > limits.max_input_bytes() {
            return invalid("tracked-change XML input exceeds the configured limit");
        }
        let map = codec::inspect_tracked_changes_source(&source, &limits)?;
        if let Some(changes) = &map.changes {
            if changes.changes.len() != map.records.len() {
                return invalid("tracked-change semantic/source record count differs");
            }
        } else if !map.records.is_empty() || map.owner.is_some() {
            return invalid("tracked-change source owner has no semantic value");
        }
        Ok(Self {
            source,
            limits,
            map: Arc::new(map),
        })
    }

    /// The exact source XML captured by this snapshot.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.source
    }

    /// Clone the shared exact source binding.
    #[must_use]
    pub fn source_arc(&self) -> Arc<str> {
        Arc::clone(&self.source)
    }

    /// The present tracked-change owner, distinguishing absence from emptiness.
    #[must_use]
    pub fn changes(&self) -> Option<&Changes> {
        self.map.changes.as_ref()
    }

    /// Presence and value of `table:track-changes` on the owner.
    #[must_use]
    pub fn tracking(&self) -> Option<bool> {
        self.map
            .owner
            .as_ref()
            .and_then(|owner| owner.tracking.as_ref())
            .and_then(|_| self.changes().map(|changes| changes.enabled))
    }

    /// Presence and value of one exact record's acceptance-state attribute.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn acceptance(&self, id: &str) -> Result<Option<Acceptance>> {
        if id.is_empty() {
            return invalid("tracked-change id must not be empty");
        }
        let index = self
            .map
            .validated
            .as_ref()
            .and_then(|validated| validated.ids.get(id))
            .copied()
            .ok_or_else(|| {
                Error::InvalidFormat(format!("tracked-change id '{id}' was not found"))
            })?;
        let record = self.map.records.get(index).ok_or_else(|| {
            Error::InvalidFormat("tracked-change source record is missing".to_string())
        })?;
        Ok(record.acceptance.as_ref().map(|_| {
            self.changes().expect("checked owner").changes[index]
                .metadata()
                .acceptance
        }))
    }

    /// Resource limits retained for edits and patch application.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Begin an inert, failure-atomic transaction with bounded fallible cloning.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn transaction(&self) -> Result<Transaction> {
        Transaction::new(self.clone())
    }

    fn acceptance_values(&self) -> &[Option<Acceptance>] {
        &self.map.acceptance
    }
}

#[derive(Clone, Debug)]
struct DraftRecord {
    source_index: Option<usize>,
    token: usize,
    dirty: bool,
    resources: Resources,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct RelationKinds(u8);

impl RelationKinds {
    fn insert(&mut self, kind: RelationKind) {
        self.0 |= 1 << kind as u8;
    }

    fn union(&mut self, other: Self) {
        self.0 |= other.0;
    }

    fn first(self) -> RelationKind {
        [
            RelationKind::Rejecting,
            RelationKind::Dependency,
            RelationKind::CellContentDeletion,
            RelationKind::ChangeDeletion,
            RelationKind::InsertionCutOff,
            RelationKind::Previous,
        ]
        .into_iter()
        .find(|kind| self.0 & (1 << *kind as u8) != 0)
        .expect("non-empty tracked-change relation mask")
    }
}

#[derive(Debug, Default)]
struct InboundBucket {
    sources: HashMap<usize, RelationKinds>,
    staged: Option<(usize, RelationKinds)>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct MultiDeletionGroup {
    start: usize,
    end: usize,
}

/// A staged inert tracked-change edit. It never applies records to live cells.
///
/// CRUD validates changed record values, configured resource totals, exact ID
/// lookup, reference existence, and required backward references before
/// publishing the operation to the draft. Dependency-changing replacements,
/// moves, and reorders check the complete graph immediately; commit rechecks
/// the complete graph before materializing source bytes.
pub struct Transaction {
    before: Snapshot,
    changes: Option<Changes>,
    tracking: Option<bool>,
    acceptance: Vec<Option<Acceptance>>,
    records: Vec<DraftRecord>,
    ids: Option<HashMap<String, usize>>,
    token_indices: HashMap<usize, usize>,
    next_source_token: usize,
    inbound: HashMap<String, InboundBucket>,
    multi_deletion_groups: Vec<MultiDeletionGroup>,
    resources: Resources,
    full_replace: bool,
}

impl Transaction {
    /// Create a transaction from a source-bound snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn new(snapshot: Snapshot) -> Result<Self> {
        let changes = match (snapshot.changes(), snapshot.map.validated.as_deref()) {
            (Some(value), Some(validated)) => {
                Some(value.try_clone_validated_with_limits(validated, &snapshot.limits)?)
            },
            (None, None) => None,
            _ => {
                return invalid("tracked-change owner and validated cache presence differ");
            },
        };
        let tracking = snapshot.tracking();
        let mut acceptance = Vec::new();
        acceptance
            .try_reserve_exact(snapshot.acceptance_values().len())
            .map_err(|_error| allocation_error("tracked-change acceptance draft"))?;
        acceptance.extend_from_slice(snapshot.acceptance_values());
        let mut records = Vec::new();
        records
            .try_reserve_exact(snapshot.map.record_resources.len())
            .map_err(|_error| allocation_error("tracked-change provenance draft"))?;
        let mut token_indices = HashMap::new();
        token_indices
            .try_reserve(snapshot.map.record_resources.len())
            .map_err(|_error| allocation_error("tracked-change source-token index"))?;
        for (source_index, resources) in snapshot.map.record_resources.iter().copied().enumerate() {
            records.push(DraftRecord {
                source_index: Some(source_index),
                token: source_index,
                dirty: false,
                resources,
            });
            token_indices.insert(source_index, source_index);
        }
        let next_source_token = records.len();
        let cached_ids = snapshot
            .map
            .validated
            .as_ref()
            .map(|validated| &validated.ids);
        let mut ids = HashMap::new();
        ids.try_reserve(cached_ids.map_or(0, HashMap::len))
            .map_err(|_error| allocation_error("tracked-change transaction ID index"))?;
        if let Some(cached_ids) = cached_ids {
            for (id, index) in cached_ids {
                ids.insert(
                    try_owned_string(id, "tracked-change transaction ID index")?,
                    *index,
                );
            }
        }
        let inbound = build_inbound_index(changes.as_ref())?;
        let multi_deletion_groups = build_multi_deletion_groups(changes.as_ref())?;
        let resources = snapshot.map.resources.unwrap_or_default();
        Ok(Self {
            before: snapshot,
            changes,
            tracking,
            acceptance,
            records,
            ids: Some(ids),
            token_indices,
            next_source_token,
            inbound,
            multi_deletion_groups,
            resources,
            full_replace: false,
        })
    }

    /// Original source snapshot used for stale checks and rollback.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Current present owner value, distinguishing absence from emptiness.
    #[must_use]
    pub fn changes(&self) -> Option<&Changes> {
        self.changes.as_ref()
    }

    /// Current presence and value of `table:track-changes`.
    #[must_use]
    pub const fn tracking(&self) -> Option<bool> {
        self.tracking
    }

    /// Set or remove the tracking attribute without accepting or applying records.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_tracking(&mut self, tracking: Option<bool>) -> Result<()> {
        let Some(changes) = self.changes.as_mut() else {
            if tracking.is_none() {
                return Ok(());
            }
            return invalid("cannot set tracking state without a tracked-changes owner");
        };
        self.tracking = tracking;
        changes.enabled = tracking.unwrap_or(false);
        Ok(())
    }

    /// Insert a validated inert record at an exact source-order index.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn insert(&mut self, index: usize, change: Change) -> Result<()> {
        let was_absent = self.changes.is_none();
        let record_count = self
            .changes
            .as_ref()
            .map_or(0, |changes| changes.changes.len());
        if index > record_count {
            return invalid(format!(
                "tracked-change insertion index {index} exceeds {record_count} records"
            ));
        }
        let record_resources = change.resources_with_limits(&self.before.limits)?;
        let base_resources = if was_absent {
            add_resources(
                self.resources,
                Resources {
                    changes: 0,
                    nodes: 1,
                    aggregate_bytes: 0,
                },
                &self.before.limits,
            )?
        } else {
            self.resources
        };
        let candidate_resources =
            add_resources(base_resources, record_resources, &self.before.limits)?;
        self.ensure_ids()?;
        if self
            .ids
            .as_ref()
            .expect("tracked-change ID index")
            .contains_key(&change.metadata().id)
        {
            return invalid(format!(
                "duplicate spreadsheet tracked-change id '{}'",
                change.metadata().id
            ));
        }
        let id = try_owned_string(&change.metadata().id, "tracked-change ID index")?;
        self.ids
            .as_mut()
            .expect("tracked-change ID index")
            .try_reserve(1)
            .map_err(|_error| allocation_error("tracked-change ID index"))?;
        let source_token = self.next_source_token;
        let next_source_token = source_token.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("tracked-change source-token overflow".to_string())
        })?;
        self.token_indices
            .try_reserve(1)
            .map_err(|_error| allocation_error("tracked-change source-token index"))?;
        reserve_one(&mut self.acceptance, "tracked-change acceptance states")?;
        reserve_one(&mut self.records, "tracked-change source records")?;
        if has_multi_deletion_marker(&change) {
            reserve_one(
                &mut self.multi_deletion_groups,
                "tracked-change multi-deletion groups",
            )?;
        }
        if was_absent {
            let mut changes = Changes::default();
            reserve_one(&mut changes.changes, "tracked-change records")?;
            self.changes = Some(changes);
            self.tracking = None;
        } else {
            reserve_one(
                &mut self.changes.as_mut().expect("existing owner").changes,
                "tracked-change records",
            )?;
        }
        if let Err(error) = self.stage_outbound(&change, source_token) {
            if was_absent {
                self.changes = None;
                self.tracking = None;
            }
            return Err(error);
        }
        let changes = self.changes.as_mut().expect("owner initialized");
        let acceptance = default_acceptance_presence(&change);
        changes.changes.insert(index, change);
        self.acceptance.insert(index, acceptance);
        self.records.insert(
            index,
            DraftRecord {
                source_index: None,
                token: source_token,
                dirty: true,
                resources: record_resources,
            },
        );
        if index < record_count {
            for value in self
                .ids
                .as_mut()
                .expect("tracked-change ID index")
                .values_mut()
            {
                if *value >= index {
                    *value += 1;
                }
            }
            for value in self.token_indices.values_mut() {
                if *value >= index {
                    *value += 1;
                }
            }
        }
        self.ids
            .as_mut()
            .expect("tracked-change ID index")
            .insert(id, index);
        self.token_indices.insert(source_token, index);
        if index < record_count {
            self.shift_inbound_for_insert(index);
        }
        self.publish_staged_outbound(index);
        let validation = self.validate_indexed_relations(
            index,
            &self.changes.as_ref().expect("owner initialized").changes[index]
                .metadata()
                .id,
        );
        let validation = validation.and_then(|()| self.validate_multi_deletion_edit(index));
        if let Err(error) = validation {
            let inserted_id = self.changes.as_ref().expect("owner initialized").changes[index]
                .metadata()
                .id
                .as_str();
            self.ids
                .as_mut()
                .expect("tracked-change ID index")
                .remove(inserted_id);
            self.token_indices.remove(&source_token);
            self.remove_outbound(index);
            self.prune_current_outbound_targets(index);
            if index < record_count {
                self.shift_inbound_for_remove(index);
            }
            if index < record_count {
                for value in self
                    .ids
                    .as_mut()
                    .expect("tracked-change ID index")
                    .values_mut()
                {
                    if *value > index {
                        *value -= 1;
                    }
                }
                for value in self.token_indices.values_mut() {
                    if *value > index {
                        *value -= 1;
                    }
                }
            }
            self.changes
                .as_mut()
                .expect("owner initialized")
                .changes
                .remove(index);
            self.acceptance.remove(index);
            self.records.remove(index);
            if was_absent {
                self.changes = None;
                self.tracking = None;
            }
            return Err(error);
        }
        self.update_multi_deletion_after_insert(index);
        self.next_source_token = next_source_token;
        self.resources = candidate_resources;
        Ok(())
    }

    /// Append a validated inert record.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn append(&mut self, change: Change) -> Result<()> {
        let index = self.changes.as_ref().map_or(0, |value| value.changes.len());
        self.insert(index, change)
    }

    /// Replace the record selected by an exact ID.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace(&mut self, id: &str, change: Change) -> Result<()> {
        let index = self.find_index(id)?;
        let record_resources = change.resources_with_limits(&self.before.limits)?;
        let candidate_resources = add_resources(
            subtract_resources(self.resources, self.records[index].resources)?,
            record_resources,
            &self.before.limits,
        )?;
        let dependencies_changed = self.changes.as_ref().expect("checked owner").changes[index]
            .metadata()
            .dependencies
            != change.metadata().dependencies;
        if change.metadata().id != id
            && self
                .ids
                .as_ref()
                .expect("tracked-change ID index")
                .contains_key(&change.metadata().id)
        {
            return invalid(format!(
                "duplicate spreadsheet tracked-change id '{}'",
                change.metadata().id
            ));
        }
        let replacement_id = try_owned_string(&change.metadata().id, "tracked-change ID index")?;
        self.ids
            .as_mut()
            .expect("tracked-change ID index")
            .try_reserve(1)
            .map_err(|_error| allocation_error("tracked-change ID index"))?;
        if has_multi_deletion_marker(&change) {
            reserve_one(
                &mut self.multi_deletion_groups,
                "tracked-change multi-deletion groups",
            )?;
        }
        if change.metadata().id != id {
            self.ensure_no_indexed_inbound(index, id)?;
        }
        let source_token = self.records[index].token;
        self.stage_outbound(&change, source_token)?;
        let before_change = mem::replace(
            &mut self.changes.as_mut().expect("checked owner").changes[index],
            change,
        );
        remove_outbound_from(&mut self.inbound, source_token, &before_change);
        let before_acceptance = self.acceptance[index];
        let before_record = self.records[index].clone();
        self.acceptance[index] = default_acceptance_presence(
            &self.changes.as_ref().expect("checked owner").changes[index],
        );
        self.records[index].dirty = true;
        self.records[index].resources = record_resources;
        let old_entry = {
            let ids = self.ids.as_mut().expect("tracked-change ID index");
            let old_entry = ids
                .remove_entry(id)
                .expect("tracked-change ID index contains selected record");
            ids.insert(replacement_id, index);
            old_entry
        };
        self.publish_staged_outbound(index);
        let replacement_id = self.changes.as_ref().expect("checked owner").changes[index]
            .metadata()
            .id
            .as_str();
        let relation = self.validate_indexed_relations(index, replacement_id);
        let relation = relation.and_then(|()| self.validate_multi_deletion_edit(index));
        let validation = relation.and_then(|()| {
            if dependencies_changed {
                self.changes
                    .as_ref()
                    .expect("checked owner")
                    .validate_graph_with_limits(&self.before.limits)
            } else {
                Ok(())
            }
        });
        if let Err(error) = validation {
            self.remove_outbound(index);
            self.stage_outbound(&before_change, source_token)
                .expect("restoring reserved tracked-change inbound edges");
            publish_staged_outbound_into(&mut self.inbound, &before_change, index);
            self.prune_current_outbound_targets(index);
            let ids = self.ids.as_mut().expect("tracked-change ID index");
            ids.remove(
                self.changes.as_ref().expect("checked owner").changes[index]
                    .metadata()
                    .id
                    .as_str(),
            );
            ids.insert(old_entry.0, index);
            self.changes.as_mut().expect("checked owner").changes[index] = before_change;
            self.acceptance[index] = before_acceptance;
            self.records[index] = before_record;
            return Err(error);
        }
        self.prune_outbound_targets(&before_change);
        self.update_multi_deletion_after_replace(index);
        self.resources = candidate_resources;
        Ok(())
    }

    /// Remove and return the record selected by an exact ID.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn remove(&mut self, id: &str) -> Result<Change> {
        let index = self.find_index(id)?;
        let shifts_following =
            index + 1 < self.changes.as_ref().expect("checked owner").changes.len();
        self.ensure_no_indexed_inbound(index, id)?;
        let candidate_resources =
            subtract_resources(self.resources, self.records[index].resources)?;
        self.remove_outbound(index);
        let changes = self.changes.as_mut().expect("checked owner");
        let change = changes.changes.remove(index);
        let acceptance = self.acceptance.remove(index);
        let record = self.records.remove(index);
        let source_token = record.token;
        let ids = self.ids.as_mut().expect("tracked-change ID index");
        let old_entry = ids
            .remove_entry(id)
            .expect("tracked-change ID index contains removed record");
        if shifts_following {
            for value in ids.values_mut() {
                if *value > index {
                    *value -= 1;
                }
            }
        }
        self.token_indices.remove(&source_token);
        if shifts_following {
            for value in self.token_indices.values_mut() {
                if *value > index {
                    *value -= 1;
                }
            }
        }
        if index < changes.changes.len() {
            self.shift_inbound_for_remove(index);
        }
        if let Err(error) = self.validate_multi_deletion_edit(index) {
            if index < self.changes.as_ref().expect("checked owner").changes.len() {
                self.shift_inbound_for_insert(index);
            }
            self.stage_outbound(&change, record.token)
                .expect("restoring reserved tracked-change inbound edges");
            publish_staged_outbound_into(&mut self.inbound, &change, index);
            self.changes
                .as_mut()
                .expect("checked owner")
                .changes
                .insert(index, change);
            self.acceptance.insert(index, acceptance);
            self.records.insert(index, record);
            let ids = self.ids.as_mut().expect("tracked-change ID index");
            if shifts_following {
                for value in ids.values_mut() {
                    if *value >= index {
                        *value += 1;
                    }
                }
                for value in self.token_indices.values_mut() {
                    if *value >= index {
                        *value += 1;
                    }
                }
            }
            ids.insert(old_entry.0, index);
            self.token_indices.insert(source_token, index);
            return Err(error);
        }
        self.prune_outbound_targets(&change);
        self.update_multi_deletion_after_remove(index);
        self.resources = candidate_resources;
        Ok(change)
    }

    /// Move one exact ID to a checked final index. `len()` means the end.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn move_to(&mut self, id: &str, index: usize) -> Result<()> {
        let from = self.find_index(id)?;
        let len = self.changes.as_ref().expect("checked owner").changes.len();
        if index > len {
            return invalid(format!(
                "tracked-change move index {index} exceeds {len} records"
            ));
        }
        let target = index.min(len.saturating_sub(1));
        if from == target {
            return Ok(());
        }
        let changes = self.changes.as_mut().expect("checked owner");
        let change = changes.changes.remove(from);
        let acceptance = self.acceptance.remove(from);
        let record = self.records.remove(from);
        changes.changes.insert(target, change);
        self.acceptance.insert(target, acceptance);
        self.records.insert(target, record);
        if let Err(error) = changes.validate_with_limits(&self.before.limits) {
            let change = changes.changes.remove(target);
            let acceptance = self.acceptance.remove(target);
            let record = self.records.remove(target);
            changes.changes.insert(from, change);
            self.acceptance.insert(from, acceptance);
            self.records.insert(from, record);
            return Err(error);
        }
        self.move_inbound_sources(from, target);
        self.rebuild_multi_deletion_groups();
        self.refresh_id_indices();
        self.refresh_token_indices();
        Ok(())
    }

    /// Replace the complete ordering with an exact permutation of current IDs.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn reorder(&mut self, ids: &[String]) -> Result<()> {
        let record_count = self
            .changes
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("tracked-changes owner is absent".to_string()))?
            .changes
            .len();
        if ids.len() != record_count {
            return invalid("tracked-change reorder must name every record exactly once");
        }
        self.ensure_ids()?;
        let mut order = Vec::new();
        order
            .try_reserve_exact(ids.len())
            .map_err(|_error| allocation_error("tracked-change reorder plan"))?;
        let mut seen = Vec::new();
        seen.try_reserve_exact(ids.len())
            .map_err(|_error| allocation_error("tracked-change reorder seen set"))?;
        seen.resize(ids.len(), false);
        for id in ids {
            let index = *self
                .ids
                .as_ref()
                .expect("tracked-change ID index")
                .get(id.as_str())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!("tracked-change id '{id}' was not found"))
                })?;
            if seen[index] {
                return invalid(format!("tracked-change id '{id}' appears more than once"));
            }
            seen[index] = true;
            order.push(index);
        }
        let mut at_position = Vec::new();
        at_position
            .try_reserve_exact(order.len())
            .map_err(|_error| allocation_error("tracked-change reorder positions"))?;
        at_position.extend(0..order.len());
        let mut position_of = Vec::new();
        position_of
            .try_reserve_exact(order.len())
            .map_err(|_error| allocation_error("tracked-change reorder inverse positions"))?;
        position_of.extend(0..order.len());

        for target in 0..order.len() {
            let desired = order[target];
            let current = position_of[desired];
            if target == current {
                continue;
            }
            self.swap_records(target, current);
            let displaced = at_position[target];
            at_position.swap(target, current);
            position_of[desired] = target;
            position_of[displaced] = current;
        }

        if let Err(error) = self
            .changes
            .as_ref()
            .expect("checked owner")
            .validate_with_limits(&self.before.limits)
        {
            for target in 0..at_position.len() {
                let current = position_of[target];
                if target == current {
                    continue;
                }
                self.swap_records(target, current);
                let displaced = at_position[target];
                at_position.swap(target, current);
                position_of[target] = target;
                position_of[displaced] = current;
            }
            return Err(error);
        }
        self.rebuild_multi_deletion_groups();
        self.refresh_id_indices();
        self.refresh_token_indices();
        Ok(())
    }

    /// Set or remove one record's acceptance-state attribute.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_acceptance(&mut self, id: &str, acceptance: Option<Acceptance>) -> Result<()> {
        let index = self.find_index(id)?;
        let before = self.changes.as_ref().expect("checked owner").changes[index]
            .metadata()
            .acceptance;
        metadata_mut(&mut self.changes.as_mut().expect("checked owner").changes[index])
            .acceptance = acceptance.unwrap_or_default();
        if let Err(error) = self.validate_indexed_relations(index, id) {
            metadata_mut(&mut self.changes.as_mut().expect("checked owner").changes[index])
                .acceptance = before;
            return Err(error);
        }
        self.acceptance[index] = acceptance;
        Ok(())
    }

    /// Remove the complete owner while leaving live spreadsheet cells untouched.
    pub fn remove_owner(&mut self) {
        self.changes = None;
        self.tracking = None;
        self.acceptance.clear();
        self.records.clear();
        self.ids = None;
        self.token_indices.clear();
        self.next_source_token = 0;
        self.inbound.clear();
        self.multi_deletion_groups.clear();
        self.resources = Resources::default();
        self.full_replace = false;
    }

    /// Explicitly replace the complete semantic owner.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn replace_all(&mut self, changes: Option<Changes>) -> Result<()> {
        let (resources, record_resources) =
            resources_for_changes(changes.as_ref(), &self.before.limits)?;
        let inbound = build_inbound_index(changes.as_ref())?;
        let multi_deletion_groups = build_multi_deletion_groups(changes.as_ref())?;
        if let Some(value) = &changes {
            value.validate_graph_with_limits(&self.before.limits)?;
        }
        let tracking = changes.as_ref().map(|value| value.enabled);
        let mut acceptance = Vec::new();
        let mut records = Vec::new();
        let mut token_indices = HashMap::new();
        if let Some(value) = &changes {
            acceptance
                .try_reserve_exact(value.changes.len())
                .map_err(|_error| allocation_error("tracked-change acceptance states"))?;
            records
                .try_reserve_exact(value.changes.len())
                .map_err(|_error| allocation_error("tracked-change source records"))?;
            token_indices
                .try_reserve(value.changes.len())
                .map_err(|_error| allocation_error("tracked-change source-token index"))?;
            for (token, (change, resources)) in
                value.changes.iter().zip(record_resources).enumerate()
            {
                acceptance.push(default_acceptance_presence(change));
                records.push(DraftRecord {
                    source_index: None,
                    token,
                    dirty: true,
                    resources,
                });
                token_indices.insert(token, token);
            }
        }
        let next_source_token = records.len();
        self.changes = changes;
        self.tracking = tracking;
        self.acceptance = acceptance;
        self.records = records;
        self.ids = None;
        self.token_indices = token_indices;
        self.next_source_token = next_source_token;
        self.inbound = inbound;
        self.multi_deletion_groups = multi_deletion_groups;
        self.resources = resources;
        self.full_replace = true;
        Ok(())
    }

    /// Restore the original semantic and presence state without reparsing.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn rollback(&mut self) -> Result<()> {
        let reset = Self::new(self.before.clone())?;
        self.changes = reset.changes;
        self.tracking = reset.tracking;
        self.acceptance = reset.acceptance;
        self.records = reset.records;
        self.ids = reset.ids;
        self.token_indices = reset.token_indices;
        self.next_source_token = reset.next_source_token;
        self.inbound = reset.inbound;
        self.multi_deletion_groups = reset.multi_deletion_groups;
        self.resources = reset.resources;
        self.full_replace = false;
        Ok(())
    }

    /// Whether the current semantic and authored-presence state differs.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changes != self.before.map.changes
            || self.tracking != self.before.tracking()
            || self.acceptance.as_slice() != self.before.acceptance_values()
    }

    /// Validate, materialize through checked descending-equivalent splices, and reparse.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        if let Some(changes) = &self.changes {
            changes.validate_with_limits(&self.before.limits)?;
        }
        if !self.is_changed() {
            let source = self.before.source_arc();
            let limits = self.before.limits;
            return Ok(Commit {
                snapshot: self.before,
                patch: Patch::new(Arc::clone(&source), source, limits),
                changed: false,
            });
        }
        let candidate = self.materialize()?;
        if candidate.len() > self.before.limits.max_output_bytes() {
            return invalid("tracked-change XML output exceeds the configured limit");
        }
        let candidate_limits = self
            .before
            .limits
            .with_max_input_bytes(self.before.limits.max_output_bytes());
        let mut snapshot =
            Snapshot::parse_with_limits(Arc::<str>::from(candidate), candidate_limits)?;
        snapshot.limits = self.before.limits.with_max_input_bytes(
            self.before
                .limits
                .max_input_bytes()
                .max(snapshot.source_xml().len()),
        );
        if snapshot.changes() != self.changes.as_ref() {
            return invalid("tracked-change candidate did not preserve requested semantic state");
        }
        if snapshot.tracking() != self.tracking {
            return invalid("tracked-change candidate did not preserve requested tracking state");
        }
        let candidate_acceptance = snapshot.acceptance_values();
        if candidate_acceptance != self.acceptance.as_slice() {
            return invalid(format!(
                "tracked-change candidate acceptance presence {candidate_acceptance:?} differs from requested {:?}",
                self.acceptance
            ));
        }
        let patch = Patch::new(
            self.before.source_arc(),
            snapshot.source_arc(),
            self.before.limits,
        );
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }

    fn materialize(&self) -> Result<String> {
        let source = self.before.source_xml();
        let owner = self.before.map.owner.as_ref();
        if self.full_replace {
            if self.changes.is_some()
                && (owner.is_some_and(|value| value.has_unsupported_content)
                    || self
                        .before
                        .map
                        .records
                        .iter()
                        .any(|record| !record.regenerable))
            {
                return invalid("cannot regenerate an unsupported tracked-change target");
            }
            return match (&self.changes, owner) {
                (None, Some(owner)) => apply_splices(
                    source,
                    vec![Splice::new(owner.element.whole, String::new())],
                    &self.before.limits,
                ),
                (None, None) => Ok(source.to_owned()),
                (Some(changes), Some(owner)) => {
                    let replacement = codec::write_tracked_changes_owner(
                        changes,
                        self.tracking,
                        &self.before.limits,
                    )?;
                    apply_splices(
                        source,
                        vec![Splice::new(owner.element.whole, replacement)],
                        &self.before.limits,
                    )
                },
                (Some(changes), None) => {
                    let replacement = codec::write_tracked_changes_owner(
                        changes,
                        self.tracking,
                        &self.before.limits,
                    )?;
                    self.insert_owner(source, replacement)
                },
            };
        }
        match (&self.changes, owner) {
            (None, Some(owner)) => apply_splices(
                source,
                vec![Splice::new(owner.element.whole, String::new())],
                &self.before.limits,
            ),
            (None, None) => Ok(source.to_owned()),
            (Some(changes), None) => {
                let replacement = codec::write_tracked_changes_owner(
                    changes,
                    self.tracking,
                    &self.before.limits,
                )?;
                self.insert_owner(source, replacement)
            },
            (Some(changes), Some(owner)) => {
                let mut splices = Vec::new();
                splices
                    .try_reserve_exact(self.before.map.records.len().saturating_add(2))
                    .map_err(|_error| allocation_error("tracked-change splice plan"))?;
                if self.tracking != self.before.tracking() {
                    let replacement = codec::rewrite_owner_tracking(
                        source,
                        owner,
                        self.tracking,
                        &self.before.limits,
                    )?;
                    let old_tail = source
                        .get(owner.element.open.end..owner.element.whole.end)
                        .ok_or_else(|| {
                            Error::InvalidFormat("tracked-change owner span is invalid".to_string())
                        })?;
                    if !replacement.ends_with(old_tail) {
                        return invalid("tracking rewrite modified tracked-change owner content");
                    }
                    let open_len =
                        replacement
                            .len()
                            .checked_sub(old_tail.len())
                            .ok_or_else(|| {
                                Error::InvalidFormat(
                                    "tracked-change owner rewrite underflow".to_string(),
                                )
                            })?;
                    splices.push(Splice::new(
                        owner.element.open,
                        replacement[..open_len].to_owned(),
                    ));
                }
                let shared = changes.changes.len().min(self.before.map.records.len());
                let original_acceptance_values = self.before.acceptance_values();
                for index in 0..shared {
                    let draft = &self.records[index];
                    let original_acceptance = draft
                        .source_index
                        .and_then(|source_index| {
                            original_acceptance_values.get(source_index).copied()
                        })
                        .flatten();
                    if draft.source_index == Some(index)
                        && !draft.dirty
                        && self.acceptance[index] == original_acceptance
                    {
                        continue;
                    }
                    let replacement = self.render_record(index)?;
                    splices.push(Splice::new(
                        self.before.map.records[index].element.whole,
                        replacement,
                    ));
                }
                for index in shared..self.before.map.records.len() {
                    splices.push(Splice::new(
                        self.before.map.records[index].element.whole,
                        String::new(),
                    ));
                }
                if changes.changes.len() > self.before.map.records.len() {
                    let mut addition = String::new();
                    for index in self.before.map.records.len()..changes.changes.len() {
                        let record = self.render_record(index)?;
                        addition.try_reserve(record.len()).map_err(|_error| {
                            allocation_error("tracked-change record insertion")
                        })?;
                        addition.push_str(&record);
                    }
                    if let Some(close) = owner.element.close {
                        splices.push(Splice::point(close.start, addition));
                    } else {
                        let intermediate = apply_splices(source, splices, &self.before.limits)?;
                        let map = codec::inspect_tracked_changes_source(
                            &intermediate,
                            &self.before.limits,
                        )?;
                        let intermediate_owner = map.owner.as_ref().ok_or_else(|| {
                            Error::InvalidFormat(
                                "tracked-change owner disappeared during rewrite".to_string(),
                            )
                        })?;
                        let replacement = codec::insert_tracked_change_into_owner(
                            &intermediate,
                            intermediate_owner,
                            &addition,
                            &self.before.limits,
                        )?;
                        return apply_splices(
                            &intermediate,
                            vec![Splice::new(intermediate_owner.element.whole, replacement)],
                            &self.before.limits,
                        );
                    }
                }
                apply_splices(source, splices, &self.before.limits)
            },
        }
    }

    fn render_record(&self, index: usize) -> Result<String> {
        let changes = self.changes.as_ref().expect("record owner");
        let draft = self.records.get(index).ok_or_else(|| {
            Error::InvalidFormat("tracked-change draft provenance is missing".to_string())
        })?;
        if let Some(source_index) = draft.source_index
            && !draft.dirty
        {
            let record = self.before.map.records.get(source_index).ok_or_else(|| {
                Error::InvalidFormat("tracked-change source provenance is invalid".to_string())
            })?;
            return codec::rewrite_record_acceptance(
                self.before.source_xml(),
                record,
                self.acceptance[index],
                &self.before.limits,
            );
        }
        if let Some(source_index) = draft.source_index {
            let record = self.before.map.records.get(source_index).ok_or_else(|| {
                Error::InvalidFormat("tracked-change source provenance is invalid".to_string())
            })?;
            if !record.regenerable || record.has_unsupported_content || record.has_rich_content {
                return invalid(format!(
                    "cannot regenerate unsupported tracked-change record '{}'",
                    record.id
                ));
            }
        }
        codec::write_tracked_change(
            &changes.changes[index],
            self.acceptance[index].is_some(),
            &self.before.limits,
        )
    }

    fn insert_owner(&self, source: &str, owner: String) -> Result<String> {
        if self.before.map.spreadsheet.close.is_some() {
            return apply_splices(
                source,
                vec![Splice::point(self.before.map.schema_insert, owner)],
                &self.before.limits,
            );
        }
        let spreadsheet = codec::insert_tracked_owner_into_spreadsheet(
            source,
            &self.before.map.spreadsheet,
            &owner,
            &self.before.limits,
        )?;
        apply_splices(
            source,
            vec![Splice::new(self.before.map.spreadsheet.whole, spreadsheet)],
            &self.before.limits,
        )
    }

    fn swap_records(&mut self, left: usize, right: usize) {
        self.changes
            .as_mut()
            .expect("tracked-change reorder owner")
            .changes
            .swap(left, right);
        self.acceptance.swap(left, right);
        self.records.swap(left, right);
    }

    fn ensure_ids(&mut self) -> Result<()> {
        if self.ids.is_some() {
            return Ok(());
        }
        let mut ids = HashMap::new();
        let count = self.changes.as_ref().map_or(0, |value| value.changes.len());
        ids.try_reserve(count)
            .map_err(|_error| allocation_error("tracked-change ID index"))?;
        if let Some(changes) = &self.changes {
            for (index, change) in changes.changes.iter().enumerate() {
                let id = try_owned_string(&change.metadata().id, "tracked-change ID index")?;
                if ids.insert(id, index).is_some() {
                    return invalid(format!(
                        "duplicate spreadsheet tracked-change id '{}'",
                        change.metadata().id
                    ));
                }
            }
        }
        self.ids = Some(ids);
        Ok(())
    }

    fn find_index(&mut self, id: &str) -> Result<usize> {
        if id.is_empty() {
            return invalid("tracked-change id must not be empty");
        }
        self.ensure_ids()?;
        self.ids
            .as_ref()
            .expect("tracked-change ID index")
            .get(id)
            .copied()
            .ok_or_else(|| Error::InvalidFormat(format!("tracked-change id '{id}' was not found")))
    }

    fn refresh_id_indices(&mut self) {
        let Some(ids) = self.ids.as_mut() else {
            return;
        };
        let Some(changes) = &self.changes else {
            ids.clear();
            return;
        };
        let mut complete = true;
        for (index, change) in changes.changes.iter().enumerate() {
            if let Some(value) = ids.get_mut(change.metadata().id.as_str()) {
                *value = index;
            } else {
                complete = false;
                break;
            }
        }
        if !complete {
            self.ids = None;
        }
    }

    fn refresh_token_indices(&mut self) {
        for (index, record) in self.records.iter().enumerate() {
            let value = self
                .token_indices
                .get_mut(&record.token)
                .expect("tracked-change source-token index");
            *value = index;
        }
    }

    fn stage_outbound(&mut self, change: &Change, source_token: usize) -> Result<()> {
        let mut failure = None;
        change.for_each_relation(|target, kind| {
            if failure.is_some() {
                return;
            }
            if !self.inbound.contains_key(target) {
                if self.inbound.try_reserve(1).is_err() {
                    failure = Some(allocation_error("tracked-change inbound index"));
                    return;
                }
                let Ok(owned_target) = try_owned_string(target, "tracked-change inbound target")
                else {
                    failure = Some(allocation_error("tracked-change inbound target"));
                    return;
                };
                self.inbound.insert(owned_target, InboundBucket::default());
            }
            let bucket = self
                .inbound
                .get_mut(target)
                .expect("tracked-change inbound target inserted");
            if bucket.sources.try_reserve(1).is_err() {
                failure = Some(allocation_error("tracked-change inbound references"));
                return;
            }
            if let Some((staged_token, kinds)) = &mut bucket.staged {
                debug_assert_eq!(*staged_token, source_token);
                kinds.insert(kind);
                return;
            }
            let mut kinds = RelationKinds::default();
            kinds.insert(kind);
            bucket.staged = Some((source_token, kinds));
        });
        if let Some(error) = failure {
            self.clear_staged_outbound(change);
            return Err(error);
        }
        Ok(())
    }

    fn clear_staged_outbound(&mut self, change: &Change) {
        change.for_each_relation(|target, _| {
            let remove = if let Some(bucket) = self.inbound.get_mut(target) {
                bucket.staged = None;
                bucket.sources.is_empty()
            } else {
                false
            };
            if remove {
                self.inbound.remove(target);
            }
        });
    }

    fn publish_staged_outbound(&mut self, source: usize) {
        let (changes, inbound) = (&self.changes, &mut self.inbound);
        let change = &changes.as_ref().expect("tracked-change owner").changes[source];
        publish_staged_outbound_into(inbound, change, source);
    }

    fn remove_outbound(&mut self, source: usize) {
        let (changes, records, inbound) = (&self.changes, &self.records, &mut self.inbound);
        let change = &changes.as_ref().expect("tracked-change owner").changes[source];
        remove_outbound_from(inbound, records[source].token, change);
    }

    fn prune_outbound_targets(&mut self, change: &Change) {
        prune_outbound_targets_from(&mut self.inbound, change);
    }

    fn prune_current_outbound_targets(&mut self, source: usize) {
        let (changes, inbound) = (&self.changes, &mut self.inbound);
        let change = &changes.as_ref().expect("tracked-change owner").changes[source];
        prune_outbound_targets_from(inbound, change);
    }

    fn shift_inbound_for_insert(&mut self, index: usize) {
        let _ = index;
    }

    fn shift_inbound_for_remove(&mut self, index: usize) {
        let _ = index;
    }

    fn move_inbound_sources(&mut self, from: usize, target: usize) {
        let _ = (from, target);
    }

    fn validate_indexed_relations(&self, selected: usize, id: &str) -> Result<()> {
        let changes = self.changes.as_ref().expect("tracked-change owner");
        let ids = self.ids.as_ref().expect("tracked-change ID index");
        changes.validate_record_relations(selected, ids)?;
        if let Some(bucket) = self.inbound.get(id) {
            for source_token in bucket.sources.keys() {
                let source = self
                    .token_indices
                    .get(source_token)
                    .copied()
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "tracked-change inbound index source is missing".to_string(),
                        )
                    })?;
                if source != selected {
                    changes.validate_record_relations(source, ids)?;
                }
            }
        }
        Ok(())
    }

    fn ensure_no_indexed_inbound(&self, selected: usize, id: &str) -> Result<()> {
        let Some(bucket) = self.inbound.get(id) else {
            return Ok(());
        };
        if let Some((source_token, kinds)) = bucket.sources.iter().find(|(source_token, _)| {
            self.token_indices
                .get(source_token)
                .is_none_or(|source| *source != selected)
        }) {
            let source = self
                .changes
                .as_ref()
                .expect("tracked-change owner")
                .changes
                .get(
                    self.token_indices
                        .get(source_token)
                        .copied()
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "tracked-change inbound index source is missing".to_string(),
                            )
                        })?,
                )
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "tracked-change inbound index source is out of bounds".to_string(),
                    )
                })?;
            return invalid(format!(
                "tracked-change id '{id}' is still referenced by '{}' ({:?})",
                source.metadata().id,
                kinds.first()
            ));
        }
        Ok(())
    }

    fn validate_multi_deletion_edit(&self, index: usize) -> Result<()> {
        let changes = self.changes.as_ref().expect("tracked-change owner");
        if changes
            .changes
            .get(index)
            .is_some_and(has_multi_deletion_marker)
        {
            changes.validate_multi_deletion_group_at(index)?;
        }
        let preceding = self
            .multi_deletion_groups
            .partition_point(|group| group.start < index);
        if let Some(group) = preceding
            .checked_sub(1)
            .and_then(|position| self.multi_deletion_groups.get(position))
            && group.end > index
        {
            changes.validate_multi_deletion_group_at(group.start)?;
        }
        Ok(())
    }

    fn update_multi_deletion_after_insert(&mut self, index: usize) {
        let span = multi_deletion_span(
            &self.changes.as_ref().expect("tracked-change owner").changes[index],
        );
        if index
            == self
                .changes
                .as_ref()
                .expect("tracked-change owner")
                .changes
                .len()
                - 1
            && span.is_none()
        {
            return;
        }
        for group in &mut self.multi_deletion_groups {
            if group.start >= index {
                group.start += 1;
                group.end += 1;
            }
        }
        if let Some(span) = span {
            let position = self
                .multi_deletion_groups
                .partition_point(|group| group.start < index);
            self.multi_deletion_groups.insert(
                position,
                MultiDeletionGroup {
                    start: index,
                    end: index + span,
                },
            );
        }
    }

    fn update_multi_deletion_after_replace(&mut self, index: usize) {
        if let Ok(position) = self
            .multi_deletion_groups
            .binary_search_by_key(&index, |group| group.start)
        {
            self.multi_deletion_groups.remove(position);
        }
        if let Some(span) = multi_deletion_span(
            &self.changes.as_ref().expect("tracked-change owner").changes[index],
        ) {
            let position = self
                .multi_deletion_groups
                .partition_point(|group| group.start < index);
            self.multi_deletion_groups.insert(
                position,
                MultiDeletionGroup {
                    start: index,
                    end: index + span,
                },
            );
        }
    }

    fn update_multi_deletion_after_remove(&mut self, index: usize) {
        if let Ok(position) = self
            .multi_deletion_groups
            .binary_search_by_key(&index, |group| group.start)
        {
            self.multi_deletion_groups.remove(position);
        }
        for group in &mut self.multi_deletion_groups {
            if group.start > index {
                group.start -= 1;
                group.end -= 1;
            }
        }
    }

    fn rebuild_multi_deletion_groups(&mut self) {
        self.multi_deletion_groups.clear();
        let Some(changes) = &self.changes else {
            return;
        };
        for (start, change) in changes.changes.iter().enumerate() {
            if let Some(span) = multi_deletion_span(change) {
                self.multi_deletion_groups.push(MultiDeletionGroup {
                    start,
                    end: start + span,
                });
            }
        }
    }
}

/// A source-exact reversible tracked-change patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source: Arc<str>,
    target: Arc<str>,
    limits: Limits,
}

impl Patch {
    fn new(source: Arc<str>, target: Arc<str>, limits: Limits) -> Self {
        Self {
            source,
            target,
            limits,
        }
    }

    /// Whether source and target bytes are identical.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source == self.target
    }

    /// Exact source XML required by this patch.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Exact target XML produced by this patch.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::new(
            Arc::clone(&self.target),
            Arc::clone(&self.source),
            self.limits,
        )
    }

    /// Apply only to the exact source snapshot from which this patch was built.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if snapshot.source_xml() != self.source() {
            return invalid("tracked-change patch source snapshot does not match");
        }
        if self.is_empty() {
            return Ok(Commit {
                snapshot: snapshot.clone(),
                patch: self.clone(),
                changed: false,
            });
        }
        if self.target.len() > self.limits.max_output_bytes() {
            return invalid("tracked-change patch target exceeds the configured output limit");
        }
        let candidate_limits = self
            .limits
            .with_max_input_bytes(self.limits.max_output_bytes());
        let mut target = Snapshot::parse_with_limits(Arc::clone(&self.target), candidate_limits)?;
        target.limits = self
            .limits
            .with_max_input_bytes(self.limits.max_input_bytes().max(target.source_xml().len()));
        Ok(Commit {
            snapshot: target,
            patch: self.clone(),
            changed: true,
        })
    }
}

/// A validated tracked-change publication candidate.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.snapshot.source_xml()
    }

    #[must_use]
    pub fn source_arc(&self) -> Arc<str> {
        self.snapshot.source_arc()
    }

    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    #[must_use]
    pub fn into_source(self) -> Arc<str> {
        self.snapshot.source
    }

    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// Apply one concise inert tracked-change edit.
///
/// # Errors
/// Returns an error when the operation cannot be completed.
pub fn update<F>(snapshot: &Snapshot, edit: F) -> Result<Commit>
where
    F: FnOnce(&mut Transaction) -> Result<()>,
{
    let mut transaction = snapshot.transaction()?;
    edit(&mut transaction)?;
    transaction.commit()
}

#[derive(Debug)]
struct Splice {
    span: codec::Span,
    replacement: String,
}

impl Splice {
    fn new(span: codec::Span, replacement: String) -> Self {
        Self { span, replacement }
    }

    fn point(at: usize, replacement: String) -> Self {
        Self::new(codec::Span { start: at, end: at }, replacement)
    }
}

fn apply_splices(source: &str, mut splices: Vec<Splice>, limits: &Limits) -> Result<String> {
    if source.len() > limits.max_input_bytes() {
        return invalid("tracked-change XML input exceeds the configured limit");
    }
    splices.sort_by_key(|splice| splice.span.start);
    let mut cursor = 0usize;
    let mut length = source.len();
    for splice in &splices {
        if splice.span.start < cursor
            || splice.span.start > splice.span.end
            || splice.span.end > source.len()
            || !source.is_char_boundary(splice.span.start)
            || !source.is_char_boundary(splice.span.end)
        {
            return invalid("tracked-change splice plan overlaps or is out of bounds");
        }
        cursor = splice.span.end;
        length = length
            .checked_sub(splice.span.end - splice.span.start)
            .and_then(|value| value.checked_add(splice.replacement.len()))
            .ok_or_else(|| {
                Error::InvalidFormat("tracked-change output size overflow".to_string())
            })?;
    }
    if length > limits.max_output_bytes() {
        return invalid("tracked-change XML output exceeds the configured limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(length)
        .map_err(|_error| allocation_error("tracked-change output"))?;
    cursor = 0;
    for splice in splices {
        output.push_str(&source[cursor..splice.span.start]);
        output.push_str(&splice.replacement);
        cursor = splice.span.end;
    }
    output.push_str(&source[cursor..]);
    Ok(output)
}

fn metadata_mut(change: &mut Change) -> &mut super::Metadata {
    match change {
        Change::Insertion(value) => &mut value.metadata,
        Change::Deletion(value) => &mut value.metadata,
        Change::Movement(value) => &mut value.metadata,
        Change::CellContent(value) => &mut value.metadata,
    }
}

fn default_acceptance_presence(change: &Change) -> Option<Acceptance> {
    match change.metadata().acceptance {
        Acceptance::Pending => None,
        value @ Acceptance::Accepted | value @ Acceptance::Rejected => Some(value),
    }
}

fn resources_for_changes(
    changes: Option<&Changes>,
    limits: &Limits,
) -> Result<(Resources, Vec<Resources>)> {
    let Some(changes) = changes else {
        return Ok((Resources::default(), Vec::new()));
    };
    let mut total = Resources {
        changes: 0,
        nodes: 1,
        aggregate_bytes: 0,
    };
    let mut records = Vec::new();
    records
        .try_reserve_exact(changes.changes.len())
        .map_err(|_error| allocation_error("tracked-change resource deltas"))?;
    for change in &changes.changes {
        let resources = change.resources_with_limits(limits)?;
        total = add_resources(total, resources, limits)?;
        records.push(resources);
    }
    Ok((total, records))
}

fn multi_deletion_span(change: &Change) -> Option<usize> {
    let Change::Deletion(deletion) = change else {
        return None;
    };
    deletion
        .multi_deletion_spanned
        .as_ref()
        .and_then(super::model::Integer::to_usize)
}

fn has_multi_deletion_marker(change: &Change) -> bool {
    matches!(change, Change::Deletion(deletion) if deletion.multi_deletion_spanned.is_some())
}

fn build_multi_deletion_groups(changes: Option<&Changes>) -> Result<Vec<MultiDeletionGroup>> {
    let mut groups = Vec::new();
    let Some(changes) = changes else {
        return Ok(groups);
    };
    for (start, change) in changes.changes.iter().enumerate() {
        let Some(span) = multi_deletion_span(change) else {
            continue;
        };
        groups
            .try_reserve(1)
            .map_err(|_error| allocation_error("tracked-change multi-deletion groups"))?;
        groups.push(MultiDeletionGroup {
            start,
            end: start.checked_add(span).ok_or_else(|| {
                Error::InvalidFormat("multi-deletion-spanned record count overflow".to_string())
            })?,
        });
    }
    Ok(groups)
}

fn publish_staged_outbound_into(
    inbound: &mut HashMap<String, InboundBucket>,
    change: &Change,
    _source: usize,
) {
    change.for_each_relation(|target, _| {
        let bucket = inbound
            .get_mut(target)
            .expect("staged tracked-change inbound target");
        let Some((source_id, kinds)) = bucket.staged.take() else {
            return;
        };
        bucket
            .sources
            .entry(source_id)
            .and_modify(|current| current.union(kinds))
            .or_insert(kinds);
    });
}

fn remove_outbound_from(
    inbound: &mut HashMap<String, InboundBucket>,
    source_token: usize,
    change: &Change,
) {
    change.for_each_relation(|target, _| {
        if let Some(bucket) = inbound.get_mut(target) {
            bucket.sources.remove(&source_token);
        }
    });
}

fn prune_outbound_targets_from(inbound: &mut HashMap<String, InboundBucket>, change: &Change) {
    change.for_each_relation(|target, _| {
        if inbound
            .get(target)
            .is_some_and(|bucket| bucket.sources.is_empty() && bucket.staged.is_none())
        {
            inbound.remove(target);
        }
    });
}

fn build_inbound_index(changes: Option<&Changes>) -> Result<HashMap<String, InboundBucket>> {
    let mut inbound = HashMap::new();
    let Some(changes) = changes else {
        return Ok(inbound);
    };
    inbound
        .try_reserve(changes.changes.len())
        .map_err(|_error| allocation_error("tracked-change inbound index"))?;
    for (source_token, change) in changes.changes.iter().enumerate() {
        let mut failure = None;
        change.for_each_relation(|target, kind| {
            if failure.is_some() {
                return;
            }
            if !inbound.contains_key(target) {
                if inbound.try_reserve(1).is_err() {
                    failure = Some(allocation_error("tracked-change inbound index"));
                    return;
                }
                let Ok(owned_target) = try_owned_string(target, "tracked-change inbound target")
                else {
                    failure = Some(allocation_error("tracked-change inbound target"));
                    return;
                };
                inbound.insert(owned_target, InboundBucket::default());
            }
            let bucket = inbound
                .get_mut(target)
                .expect("tracked-change inbound target inserted");
            if bucket.sources.try_reserve(1).is_err() {
                failure = Some(allocation_error("tracked-change inbound references"));
                return;
            }
            if let Some(kinds) = bucket.sources.get_mut(&source_token) {
                kinds.insert(kind);
                return;
            }
            let mut kinds = RelationKinds::default();
            kinds.insert(kind);
            bucket.sources.insert(source_token, kinds);
        });
        if let Some(error) = failure {
            return Err(error);
        }
    }
    Ok(inbound)
}

fn add_resources(left: Resources, right: Resources, limits: &Limits) -> Result<Resources> {
    let resources = Resources {
        changes: left
            .changes
            .checked_add(right.changes)
            .ok_or_else(|| Error::InvalidFormat("tracked-change count overflow".to_string()))?,
        nodes: left.nodes.checked_add(right.nodes).ok_or_else(|| {
            Error::InvalidFormat("tracked-change node count overflow".to_string())
        })?,
        aggregate_bytes: left
            .aggregate_bytes
            .checked_add(right.aggregate_bytes)
            .ok_or_else(|| {
                Error::InvalidFormat("tracked-change aggregate size overflow".to_string())
            })?,
    };
    if resources.changes > limits.max_changes() {
        return invalid("spreadsheet tracked-change count exceeds resource limit");
    }
    if resources.nodes > limits.max_nodes() {
        return invalid("spreadsheet tracked-change node count exceeds resource limit");
    }
    if resources.aggregate_bytes > limits.max_aggregate_bytes() {
        return invalid("tracked-change metadata exceeds configured aggregate limit");
    }
    Ok(resources)
}

fn subtract_resources(left: Resources, right: Resources) -> Result<Resources> {
    Ok(Resources {
        changes: left
            .changes
            .checked_sub(right.changes)
            .ok_or_else(|| Error::InvalidFormat("tracked-change count underflow".to_string()))?,
        nodes: left.nodes.checked_sub(right.nodes).ok_or_else(|| {
            Error::InvalidFormat("tracked-change node count underflow".to_string())
        })?,
        aggregate_bytes: left
            .aggregate_bytes
            .checked_sub(right.aggregate_bytes)
            .ok_or_else(|| {
                Error::InvalidFormat("tracked-change aggregate size underflow".to_string())
            })?,
    })
}

fn reserve_one<T>(values: &mut Vec<T>, label: &str) -> Result<()> {
    values
        .try_reserve(1)
        .map_err(|_error| allocation_error(label))
}

#[cfg(test)]
mod cache_regressions {
    use super::super::model::{
        Deletion, Dimension, Info, Insertion, Metadata, NestedDeletion, PositiveInteger,
    };
    use super::*;

    fn metadata(id: String, dependencies: Vec<String>) -> Metadata {
        Metadata {
            id,
            acceptance: Acceptance::Pending,
            rejecting_change_id: None,
            info: Info {
                creator: Some("A".to_string()),
                date: Some("2026-08-08T00:00:00Z".to_string()),
                comments: Vec::new(),
            },
            dependencies,
            deletions: Vec::new(),
        }
    }

    fn insertion(id: String, dependencies: Vec<String>) -> Change {
        Change::Insertion(Insertion {
            metadata: metadata(id, dependencies),
            dimension: Dimension::Row,
            position: 0.into(),
            count: PositiveInteger::try_from(1usize)
                .expect("test fixture or operation should succeed"),
            table: None,
        })
    }

    fn deletion(id: &str, span: Option<i64>) -> Change {
        Change::Deletion(Deletion {
            metadata: metadata(id.to_string(), Vec::new()),
            dimension: Dimension::Row,
            position: 0.into(),
            table: None,
            multi_deletion_spanned: span.map(Into::into),
            cut_offs: Vec::new(),
        })
    }

    fn empty_transaction() -> Transaction {
        Snapshot::parse(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:tracked-changes/></office:spreadsheet></office:body></office:document-content>"#,
        )
        .expect("test fixture or operation should succeed")
        .transaction()
        .expect("test fixture or operation should succeed")
    }

    #[test]
    fn rejected_targets_and_completed_groups_keep_caches_bounded() {
        let mut transaction = empty_transaction();
        for index in 0..128 {
            let missing = format!("missing-{index}");
            let candidate = insertion(format!("rejected-{index}"), vec![missing]);
            assert!(transaction.append(candidate).is_err());
            assert!(transaction.inbound.is_empty());
            assert!(
                transaction
                    .changes()
                    .expect("test fixture or operation should succeed")
                    .changes
                    .is_empty()
            );
        }

        transaction
            .append(insertion("base".to_string(), Vec::new()))
            .expect("test fixture or operation should succeed");
        for index in 0..128 {
            let missing = format!("replacement-missing-{index}");
            let candidate = insertion("base".to_string(), vec![missing]);
            assert!(transaction.replace("base", candidate).is_err());
            assert!(transaction.inbound.is_empty());
            assert_eq!(
                transaction
                    .changes()
                    .expect("test fixture or operation should succeed")
                    .changes
                    .len(),
                1
            );
        }
        transaction
            .remove("base")
            .expect("test fixture or operation should succeed");

        transaction
            .append(insertion("hot".to_string(), Vec::new()))
            .expect("test fixture or operation should succeed");
        for index in 0..128 {
            let id = format!("hot-source-{index}");
            let mut candidate = insertion(id, vec!["hot".to_string()]);
            let Change::Insertion(value) = &mut candidate else {
                unreachable!();
            };
            for _ in 0..16 {
                value.metadata.deletions.push(NestedDeletion::Change {
                    change_id: Some("hot".to_string()),
                });
            }
            transaction
                .append(candidate)
                .expect("test fixture or operation should succeed");
        }
        assert_eq!(transaction.inbound["hot"].sources.len(), 128);
        for index in (0..128).rev() {
            transaction
                .remove(&format!("hot-source-{index}"))
                .expect("test fixture or operation should succeed");
        }
        assert!(!transaction.inbound.contains_key("hot"));
        transaction
            .remove("hot")
            .expect("test fixture or operation should succeed");

        let targets: Vec<String> = (0..64).map(|index| format!("target-{index}")).collect();
        for target in &targets {
            transaction
                .append(insertion(target.clone(), Vec::new()))
                .expect("test fixture or operation should succeed");
        }
        let long_source_id = "source".repeat(512);
        transaction
            .append(insertion(long_source_id.clone(), targets.clone()))
            .expect("test fixture or operation should succeed");
        let source_token = transaction
            .records
            .last()
            .expect("test fixture or operation should succeed")
            .token;
        for target in &targets {
            let bucket = &transaction.inbound[target];
            assert_eq!(bucket.sources.len(), 1);
            assert!(bucket.sources.contains_key(&source_token));
        }
        transaction
            .remove(&long_source_id)
            .expect("test fixture or operation should succeed");
        for target in targets.into_iter().rev() {
            transaction
                .remove(&target)
                .expect("test fixture or operation should succeed");
        }
        assert!(transaction.inbound.is_empty());

        transaction
            .replace_all(Some(Changes {
                enabled: false,
                changes: vec![deletion("d1", Some(2)), deletion("d2", None)],
            }))
            .expect("test fixture or operation should succeed");
        let completed = transaction.multi_deletion_groups.clone();
        for index in 0..128 {
            transaction
                .append(insertion(format!("valid-{index}"), Vec::new()))
                .expect("test fixture or operation should succeed");
            assert_eq!(transaction.multi_deletion_groups, completed);
        }
        assert!(transaction.inbound.is_empty());
    }
}

fn try_owned_string(value: &str, label: &str) -> Result<String> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| allocation_error(label))?;
    owned.push_str(value);
    Ok(owned)
}

fn allocation_error(label: &str) -> Error {
    Error::InvalidFormat(format!("unable to allocate bounded {label}"))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
