//! Isolated cell transactions, disjoint joins, and source-checked patches.

use std::collections::{BTreeMap, btree_map::Entry};
use std::fmt;
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, Relationship};
use litchi_sheet::{At, Cell as Address};

use super::{Sheet, SheetKind, SheetSelector, Workbook};
use crate::cell::{Cell, Content, Stored};
use crate::error::{EditBlock, Error, Result, invalid};
use crate::raw;
use crate::raw::worksheet::edit::{Action, Payload, StyleEffect};
use crate::{Style, StyleKey, StyleState};

/// Cell state recorded before or after one semantic change.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum State {
    Missing,
    Cell { content: Cell, style: StyleState },
}

impl State {
    fn read(value: Option<&Stored>, workbook: &Workbook) -> Self {
        value.map_or(Self::Missing, |stored| Self::Cell {
            content: stored.cell.clone(),
            style: stored.style.map_or(StyleState::Default, |key| {
                StyleState::Shared(StyleKey::new(
                    key,
                    Arc::clone(&workbook.inner.style_lineage),
                ))
            }),
        })
    }

    fn after(before: Option<&Stored>, action: &Action, workbook: &Workbook) -> Self {
        let Action::Update { payload, style } = action else {
            return Self::Missing;
        };
        let exists = before.is_some() || action.creates_missing();
        if !exists {
            return Self::Missing;
        }
        let content = match payload {
            Some(Payload::Set(content)) => content.as_cell(),
            Some(Payload::Clear | Payload::ClearIfPresent) => Cell::Empty,
            None => before.map_or(Cell::Empty, |stored| stored.cell.clone()),
        };
        let style = match style {
            Some(StyleEffect::Set(key)) => StyleState::Shared(StyleKey::new(
                *key,
                Arc::clone(&workbook.inner.style_lineage),
            )),
            Some(StyleEffect::Reset) => StyleState::Default,
            None => before
                .and_then(|stored| stored.style)
                .map_or(StyleState::Default, |key| {
                    StyleState::Shared(StyleKey::new(
                        key,
                        Arc::clone(&workbook.inner.style_lineage),
                    ))
                }),
        };
        Self::Cell { content, style }
    }

    fn rebind_style(&mut self, workbook: &Workbook) {
        if let Self::Cell { style, .. } = self {
            style.rebind(&workbook.inner.style_lineage);
        }
    }

    const fn uses_shared_style(&self) -> bool {
        matches!(
            self,
            Self::Cell {
                style: StyleState::Shared(_),
                ..
            }
        )
    }
}

/// One deterministic semantic cell change in a patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    sheet: Box<str>,
    address: Address,
    before: State,
    after: State,
}

impl Change {
    pub fn sheet(&self) -> &str {
        &self.sheet
    }

    pub fn address(&self) -> Address {
        self.address
    }

    pub fn before(&self) -> &State {
        &self.before
    }

    pub fn after(&self) -> &State {
        &self.after
    }
}

/// Overlapping cell effects on one logical worksheet.
#[derive(Debug, PartialEq, Eq)]
pub struct Conflict {
    sheet: Box<str>,
    position: usize,
    addresses: Box<[Address]>,
}

impl Conflict {
    /// Developer-facing worksheet name.
    pub fn sheet(&self) -> &str {
        &self.sheet
    }

    /// Checked zero-based worksheet position in the shared base snapshot.
    pub fn position(&self) -> usize {
        self.position
    }

    /// Deterministically ordered cells written by both edits.
    pub fn addresses(&self) -> &[Address] {
        &self.addresses
    }
}

/// Structured overlap report returned by [`Edit::join`].
#[derive(Debug, PartialEq, Eq)]
pub struct ConflictSet {
    conflicts: Box<[Conflict]>,
}

impl ConflictSet {
    /// Conflicts grouped in worksheet order.
    pub fn conflicts(&self) -> &[Conflict] {
        &self.conflicts
    }

    /// Number of overlapping cell effects across all worksheets.
    pub fn len(&self) -> usize {
        self.conflicts
            .iter()
            .map(|conflict| conflict.addresses.len())
            .sum()
    }

    /// Whether no overlapping cell effects were found.
    pub fn is_empty(&self) -> bool {
        self.conflicts.is_empty()
    }
}

impl fmt::Display for ConflictSet {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{} overlapping cell effect(s) across {} worksheet(s)",
            self.len(),
            self.conflicts.len()
        )
    }
}

/// Why two independently prepared edits could not be joined.
#[derive(Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum JoinFailure {
    /// The edits were not created from the same immutable snapshot lineage.
    DifferentSnapshot,
    /// Both edits write at least one of the same cells.
    Overlap(ConflictSet),
}

/// Recoverable join failure that returns ownership of the rejected edit.
pub struct JoinError {
    failure: JoinFailure,
    rejected: Edit,
}

impl fmt::Debug for JoinError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("JoinError")
            .field("failure", &self.failure)
            .field("rejected_effects", &self.rejected.len())
            .finish()
    }
}

impl JoinError {
    /// Structured reason the join was refused.
    pub fn failure(&self) -> &JoinFailure {
        &self.failure
    }

    /// Overlapping effects, or `None` for a lineage mismatch.
    pub fn conflicts(&self) -> Option<&ConflictSet> {
        match &self.failure {
            JoinFailure::Overlap(conflicts) => Some(conflicts),
            JoinFailure::DifferentSnapshot => None,
        }
    }

    /// Borrow the edit that was not merged.
    pub fn rejected(&self) -> &Edit {
        &self.rejected
    }

    /// Recover the edit that was not merged.
    pub fn into_rejected(self) -> Edit {
        self.rejected
    }

    /// Recover both the structured reason and rejected edit.
    pub fn into_parts(self) -> (JoinFailure, Edit) {
        (self.failure, self.rejected)
    }
}

impl fmt::Display for JoinError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.failure {
            JoinFailure::DifferentSnapshot => {
                formatter.write_str("edits belong to different workbook snapshots")
            },
            JoinFailure::Overlap(conflicts) => {
                write!(formatter, "edit effects overlap: {conflicts}")
            },
        }
    }
}

impl std::error::Error for JoinError {}

#[derive(Debug, Clone)]
struct PartChange {
    uri: PackURI,
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
}

#[derive(Debug, Clone)]
struct StyleGuard {
    uri: PackURI,
    content: Arc<Vec<u8>>,
}

impl StyleGuard {
    fn validate(&self, workbook: &Workbook) -> Result<()> {
        let matches = workbook.inner.styles_uri.as_ref() == Some(&self.uri)
            && workbook
                .inner
                .package
                .get_part(&self.uri)
                .is_ok_and(|part| part.blob() == self.content.as_slice());
        if matches {
            Ok(())
        } else {
            Err(Error::PatchConflict {
                part: self.uri.to_string(),
            })
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum GraphAction {
    Add,
    Remove,
}

#[derive(Clone)]
struct GraphChange {
    action: GraphAction,
    source: PackURI,
    relationship: Relationship,
    part: Box<dyn Part + Send + Sync>,
}

impl std::fmt::Debug for GraphChange {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("GraphChange")
            .field("action", &self.action)
            .field("source", &self.source)
            .field("relationship", &self.relationship.r_id())
            .field("part", self.part.partname())
            .finish()
    }
}

/// Versioned reversible patch produced by a successful transaction.
///
/// The current representation is an in-memory source-checked patch. Its public
/// changes are semantic and contain no native Office IDs; private deltas retain
/// exact changed-part bytes and relationship fields until the deterministic
/// wire format lands.
#[derive(Debug, Clone, Default)]
pub struct Patch {
    changes: Box<[Change]>,
    parts: Box<[PartChange]>,
    graph: Box<[GraphChange]>,
    style_guard: Option<StyleGuard>,
}

impl Patch {
    pub const VERSION: u16 = 1;

    pub const fn version(&self) -> u16 {
        Self::VERSION
    }

    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    pub fn len(&self) -> usize {
        self.changes.len()
    }

    pub fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// Build the inverse without copying part payloads.
    pub fn inverse(&self) -> Self {
        Self {
            changes: self
                .changes
                .iter()
                .map(|change| Change {
                    sheet: change.sheet.clone(),
                    address: change.address,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            parts: self
                .parts
                .iter()
                .map(|part| PartChange {
                    uri: part.uri.clone(),
                    before: Arc::clone(&part.after),
                    after: Arc::clone(&part.before),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            graph: self
                .graph
                .iter()
                .map(|change| GraphChange {
                    action: match change.action {
                        GraphAction::Add => GraphAction::Remove,
                        GraphAction::Remove => GraphAction::Add,
                    },
                    source: change.source.clone(),
                    relationship: change.relationship.clone(),
                    part: change.part.clone(),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            style_guard: self.style_guard.clone(),
        }
    }

    pub(super) fn apply_to(&self, workbook: &Workbook) -> Result<Commit> {
        ensure_unsigned(workbook)?;
        if let Some(guard) = &self.style_guard {
            guard.validate(workbook)?;
        }
        if self.parts.is_empty() && self.graph.is_empty() {
            return Ok(Commit {
                workbook: workbook.clone(),
                patch: self.clone(),
            });
        }
        let mut package = workbook.inner.package.clone();
        for change in &self.parts {
            let part = package.get_part(&change.uri)?;
            if part.blob() != change.before.as_slice() {
                return Err(Error::PatchConflict {
                    part: change.uri.to_string(),
                });
            }
        }
        for change in &self.graph {
            change.validate(&package)?;
        }
        for change in &self.parts {
            package
                .get_part_mut(&change.uri)?
                .set_blob_shared(Arc::clone(&change.after));
        }
        for change in &self.graph {
            change.apply(&mut package)?;
        }
        let workbook = Workbook::from_package_with_styles(package, Some(workbook))?;
        let mut patch = self.clone();
        for change in &mut patch.changes {
            change.before.rebind_style(&workbook);
            change.after.rebind_style(&workbook);
        }
        Ok(Commit { workbook, patch })
    }
}

/// Successful atomic transaction result.
#[derive(Debug)]
pub struct Commit {
    workbook: Workbook,
    patch: Patch,
}

impl Commit {
    pub fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    pub fn into_workbook(self) -> Workbook {
        self.workbook
    }

    pub fn into_parts(self) -> (Workbook, Patch) {
        (self.workbook, self.patch)
    }
}

impl GraphChange {
    fn validate(&self, package: &OpcPackage) -> Result<()> {
        let source = package.get_part(&self.source)?;
        match self.action {
            GraphAction::Remove => {
                let relationship =
                    source.rels().get(self.relationship.r_id()).ok_or_else(|| {
                        Error::PatchConflict {
                            part: self.source.to_string(),
                        }
                    })?;
                if !same_relationship(relationship, &self.relationship) {
                    return Err(Error::PatchConflict {
                        part: self.source.to_string(),
                    });
                }
                let part = package.get_part(self.part.partname())?;
                if !same_part(part, &*self.part) {
                    return Err(Error::PatchConflict {
                        part: self.part.partname().to_string(),
                    });
                }
            },
            GraphAction::Add => {
                if source.rels().get(self.relationship.r_id()).is_some()
                    || package
                        .validate_new_part_name(self.part.partname())
                        .is_err()
                {
                    return Err(Error::PatchConflict {
                        part: self.part.partname().to_string(),
                    });
                }
            },
        }
        Ok(())
    }

    fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        match self.action {
            GraphAction::Remove => {
                package
                    .get_part_mut(&self.source)?
                    .rels_mut()
                    .remove(self.relationship.r_id())
                    .ok_or_else(|| Error::PatchConflict {
                        part: self.source.to_string(),
                    })?;
                if !package.remove_part(self.part.partname()) {
                    return Err(Error::PatchConflict {
                        part: self.part.partname().to_string(),
                    });
                }
            },
            GraphAction::Add => {
                package.try_add_part(self.part.clone())?;
                package
                    .get_part_mut(&self.source)?
                    .rels_mut()
                    .try_add_relationship(
                        self.relationship.reltype().to_owned(),
                        self.relationship.target_ref().to_owned(),
                        self.relationship.r_id().to_owned(),
                        self.relationship.target_mode(),
                    )?;
            },
        }
        Ok(())
    }
}

/// Isolated workbook transaction. Dropping it rolls back every pending change.
#[derive(Debug)]
pub struct Edit {
    base: Workbook,
    sheets: BTreeMap<usize, BTreeMap<Address, Action>>,
}

impl Edit {
    pub(super) fn new(base: Workbook) -> Result<Self> {
        ensure_unsigned(&base)?;
        Ok(Self {
            base,
            sheets: BTreeMap::new(),
        })
    }

    /// Select a worksheet for short transaction-scoped cell operations.
    pub fn sheet<'e, 's>(
        &'e mut self,
        selector: impl Into<SheetSelector<'s>>,
    ) -> Result<Option<SheetEdit<'e>>> {
        let sheet = self.base.sheet(selector)?;
        let Some(sheet) = sheet else {
            return Ok(None);
        };
        if sheet.kind() != SheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name().to_owned(),
            });
        }
        let position = sheet.position();
        Ok(Some(SheetEdit {
            edit: self,
            position,
        }))
    }

    pub fn len(&self) -> usize {
        self.sheets.values().map(BTreeMap::len).sum()
    }

    pub fn is_empty(&self) -> bool {
        self.sheets.values().all(BTreeMap::is_empty)
    }

    /// Join an independently prepared edit when every cell effect is disjoint.
    ///
    /// Both edits must originate from the same immutable snapshot. On failure
    /// `self` is unchanged and [`JoinError`] returns ownership of `other`, so
    /// callers never lose prepared work while resolving a conflict.
    pub fn join(&mut self, other: Self) -> std::result::Result<&mut Self, JoinError> {
        if !Arc::ptr_eq(&self.base.inner, &other.base.inner) {
            return Err(JoinError {
                failure: JoinFailure::DifferentSnapshot,
                rejected: other,
            });
        }
        let conflicts = self.conflicts_with(&other);
        if !conflicts.is_empty() {
            return Err(JoinError {
                failure: JoinFailure::Overlap(conflicts),
                rejected: other,
            });
        }

        for (position, actions) in other.sheets {
            match self.sheets.entry(position) {
                Entry::Vacant(entry) => {
                    entry.insert(actions);
                },
                Entry::Occupied(entry) => {
                    let accepted = entry.into_mut();
                    for (address, action) in actions {
                        match accepted.entry(address) {
                            Entry::Vacant(entry) => {
                                entry.insert(action);
                            },
                            Entry::Occupied(mut entry) => {
                                entry.get_mut().merge(action);
                            },
                        }
                    }
                },
            }
        }
        Ok(self)
    }

    /// Validate and atomically publish a new immutable snapshot.
    pub fn commit(self) -> Result<Commit> {
        ensure_unsigned(&self.base)?;
        if self.is_empty() {
            return Ok(Commit {
                workbook: self.base,
                patch: Patch::default(),
            });
        }
        let Self { base, sheets } = self;
        let mut changes = Vec::new();
        let mut parts = Vec::new();
        let graph = calculation_chain_removal(&base)?;
        for (position, requested) in sheets {
            let data =
                base.inner.sheets.get(position).cloned().ok_or_else(|| {
                    invalid(format!("edited sheet position {position} disappeared"))
                })?;
            if data.kind != SheetKind::Worksheet {
                return Err(Error::NotWorksheet {
                    sheet: data.name.clone(),
                });
            }
            let sheet = Sheet {
                owner: Arc::clone(&base.inner),
                data: Arc::clone(&data),
            };
            let store = sheet.store()?;
            let mut effective = BTreeMap::new();
            for (address, action) in requested {
                let before = store.entry(address);
                if before.is_some_and(|stored| matches!(stored.cell, Cell::Unknown(_)))
                    && action.payload().is_some()
                {
                    return Err(Error::EditBlocked {
                        sheet: data.name.clone(),
                        address,
                        reason: EditBlock::UnknownCell,
                    });
                }
                let before_state = State::read(before, &base);
                let after_state = State::after(before, &action, &base);
                if before_state == after_state {
                    continue;
                }
                effective.insert(address, action);
                changes.push(Change {
                    sheet: data.name.clone().into_boxed_str(),
                    address,
                    before: before_state,
                    after: after_state,
                });
            }
            if effective.is_empty() {
                continue;
            }

            let part = base.inner.package.get_part(&data.part_uri)?;
            let before = part.blob_arc();
            let after = raw::worksheet::edit::rewrite(&before, &data.name, &effective)?;
            let parsed = raw::worksheet::parse(&after, || base.inner.shared_strings())?;
            base.inner.validate_styles(&parsed)?;
            for change in changes.iter().rev().take(effective.len()) {
                let actual = State::read(parsed.entry(change.address), &base);
                if actual != change.after {
                    return Err(invalid(format!(
                        "worksheet edit verification failed at {}!{}",
                        change.sheet, change.address
                    )));
                }
            }
            parts.push(PartChange {
                uri: data.part_uri.clone(),
                before,
                after: Arc::new(after),
            });
        }

        if parts.is_empty() {
            return Ok(Commit {
                workbook: base,
                patch: Patch::default(),
            });
        }
        let style_guard = changes
            .iter()
            .any(|change| change.before.uses_shared_style() || change.after.uses_shared_style())
            .then(|| {
                let uri = base
                    .inner
                    .styles_uri
                    .as_ref()
                    .ok_or_else(|| invalid("shared style state has no styles part"))?
                    .clone();
                let content = base.inner.package.get_part(&uri)?.blob_arc();
                Ok::<_, Error>(StyleGuard { uri, content })
            })
            .transpose()?;
        let workbook_part = base.inner.package.get_part(&base.inner.workbook_uri)?;
        let before = workbook_part.blob_arc();
        let after = raw::recalc::invalidate(&before)?;
        if after.as_slice() != before.as_slice() {
            parts.push(PartChange {
                uri: base.inner.workbook_uri.clone(),
                before,
                after: Arc::new(after),
            });
        }
        let mut package = base.inner.package.clone();
        for part in &parts {
            package
                .get_part_mut(&part.uri)?
                .set_blob_shared(Arc::clone(&part.after));
        }
        for change in &graph {
            change.validate(&package)?;
            change.apply(&mut package)?;
        }
        let workbook = Workbook::from_package_with_styles(package, Some(&base))?;
        Ok(Commit {
            workbook,
            patch: Patch {
                changes: changes.into_boxed_slice(),
                parts: parts.into_boxed_slice(),
                graph: graph.into_boxed_slice(),
                style_guard,
            },
        })
    }

    fn actions(&mut self, position: usize) -> &mut BTreeMap<Address, Action> {
        self.sheets.entry(position).or_default()
    }

    fn conflicts_with(&self, other: &Self) -> ConflictSet {
        let mut conflicts = Vec::new();
        for (position, left) in &self.sheets {
            let Some(right) = other.sheets.get(position) else {
                continue;
            };
            let mut addresses = Vec::new();
            let mut left = left.iter().peekable();
            let mut right = right.iter().peekable();
            while let (Some((left_address, left_action)), Some((right_address, right_action))) =
                (left.peek(), right.peek())
            {
                match left_address.cmp(right_address) {
                    std::cmp::Ordering::Less => {
                        left.next();
                    },
                    std::cmp::Ordering::Greater => {
                        right.next();
                    },
                    std::cmp::Ordering::Equal => {
                        if left_action.overlaps(right_action) {
                            addresses.push(**left_address);
                        }
                        left.next();
                        right.next();
                    },
                }
            }
            if addresses.is_empty() {
                continue;
            }
            let sheet = self
                .base
                .inner
                .sheets
                .get(*position)
                .map_or("<missing sheet>", |sheet| sheet.name.as_str());
            conflicts.push(Conflict {
                sheet: sheet.into(),
                position: *position,
                addresses: addresses.into_boxed_slice(),
            });
        }
        ConflictSet {
            conflicts: conflicts.into_boxed_slice(),
        }
    }
}

/// Borrowed worksheet editor tied to one transaction.
#[derive(Debug)]
pub struct SheetEdit<'a> {
    edit: &'a mut Edit,
    position: usize,
}

impl SheetEdit<'_> {
    /// Create or replace a cell's primary value or formula.
    pub fn set<'a>(
        &mut self,
        at: impl Into<At<'a>>,
        content: impl Into<Content>,
    ) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let content = content.into();
        content.validate_for_write()?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::set(content));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(Payload::Set(content)),
        }
        Ok(self)
    }

    /// Remove primary value/formula content while retaining the cell record,
    /// local style, metadata, comments, and unknown non-payload children.
    pub fn clear<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let actions = self.edit.actions(self.position);
        let create = actions.get(&address).is_some_and(|action| {
            matches!(action.payload(), Some(Payload::Set(_) | Payload::Clear))
        });
        match actions.entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::clear(create));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_payload(if create {
                Payload::Clear
            } else {
                Payload::ClearIfPresent
            }),
        }
        Ok(self)
    }

    /// Apply an existing shared style without copying its formatting payload.
    ///
    /// The handle must belong to this edit's shared-style table lineage.
    /// Styling a missing cell creates an explicit empty cell record.
    pub fn style<'a>(&mut self, at: impl Into<At<'a>>, style: &Style) -> Result<&mut Self> {
        if !Arc::ptr_eq(
            &self.edit.base.inner.style_lineage,
            &style.owner.style_lineage,
        ) {
            return Err(Error::ForeignStyle);
        }
        let address = at.into().resolve()?;
        let effect = StyleEffect::Set(style.raw());
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::style(style.raw()));
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(effect),
        }
        Ok(self)
    }

    /// Remove an explicit local style reference, retaining the cell payload.
    ///
    /// A missing cell remains missing and resolves as a no-op at commit.
    pub fn reset_style<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        match self.edit.actions(self.position).entry(address) {
            Entry::Vacant(entry) => {
                entry.insert(Action::reset_style());
            },
            Entry::Occupied(mut entry) => entry.get_mut().set_style(StyleEffect::Reset),
        }
        Ok(self)
    }

    /// Remove the cell record without shifting surrounding cells.
    pub fn remove<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        self.edit
            .actions(self.position)
            .insert(address, Action::Remove);
        Ok(self)
    }
}

fn ensure_unsigned(workbook: &Workbook) -> Result<()> {
    if workbook.inner.package.has_digital_signatures() {
        Err(Error::Signed)
    } else {
        Ok(())
    }
}

fn calculation_chain_removal(workbook: &Workbook) -> Result<Vec<GraphChange>> {
    let main = workbook
        .inner
        .package
        .get_part(&workbook.inner.workbook_uri)?;
    let mut matching = main.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::CALC_CHAIN
                | litchi_opc::constants::relationship_type::STRICT_CALC_CHAIN
        )
    });
    let Some(relationship) = matching.next() else {
        return Ok(Vec::new());
    };
    if matching.next().is_some() {
        return Err(invalid(
            "workbook has multiple calculation-chain relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("calculation-chain relationship cannot be external"));
    }
    let target = relationship.target_partname()?;
    let part = workbook.inner.package.get_part(&target)?;
    ensure_exclusive_incoming_relationship(
        &workbook.inner.package,
        part.partname(),
        &workbook.inner.workbook_uri,
        relationship.r_id(),
    )?;
    Ok(vec![GraphChange {
        action: GraphAction::Remove,
        source: workbook.inner.workbook_uri.clone(),
        relationship: relationship.clone(),
        part: part.clone_part(),
    }])
}

fn ensure_exclusive_incoming_relationship(
    package: &OpcPackage,
    target: &PackURI,
    expected_source: &PackURI,
    expected_id: &str,
) -> Result<()> {
    let targets = |relationship: &Relationship| -> Result<bool> {
        if relationship.is_external() {
            return Ok(false);
        }
        relationship
            .target_partname()
            .map(|candidate| candidate.as_str().eq_ignore_ascii_case(target.as_str()))
            .map_err(Into::into)
    };
    for relationship in package.rels().iter() {
        if targets(relationship)? {
            return Err(invalid(format!(
                "calculation-chain part '{target}' has another incoming relationship"
            )));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if targets(relationship)?
                && !(source.partname() == expected_source && relationship.r_id() == expected_id)
            {
                return Err(invalid(format!(
                    "calculation-chain part '{target}' has another incoming relationship"
                )));
            }
        }
    }
    Ok(())
}

fn same_relationship(left: &Relationship, right: &Relationship) -> bool {
    left.r_id() == right.r_id()
        && left.reltype() == right.reltype()
        && left.target_ref() == right.target_ref()
        && left.target_mode() == right.target_mode()
}

fn same_part(left: &dyn Part, right: &dyn Part) -> bool {
    left.partname() == right.partname()
        && left.content_type() == right.content_type()
        && left.blob() == right.blob()
        && left.rels().len() == right.rels().len()
        && left.rels().iter().all(|relationship| {
            right
                .rels()
                .get(relationship.r_id())
                .is_some_and(|other| same_relationship(relationship, other))
        })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::{Number, Value};
    use crate::formula::Formula;
    use litchi_opc::{BlobPart, TargetMode};

    #[test]
    fn cell_crud_is_atomic_reversible_and_source_preserving() {
        let source = Workbook::new().expect("source workbook");
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        {
            let mut sheet = edit.sheet("Sheet1").expect("sheet lookup").expect("sheet");
            sheet
                .set("A1", "hello")
                .and_then(|sheet| sheet.set("B2", 42_i32))
                .and_then(|sheet| sheet.set("C3", Formula::new("B2*2").expect("formula")))
                .expect("cell changes");
        }
        let committed = edit.commit().expect("commit");
        assert_eq!(
            source.to_bytes().expect("source remains valid"),
            source_bytes
        );
        assert_eq!(committed.patch().len(), 3);

        let book = committed.workbook();
        let sheet = book.sheet("Sheet1").expect("lookup").expect("sheet");
        assert!(matches!(
            sheet.cell("A1").expect("A1"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "hello"
        ));
        assert!(matches!(
            sheet.cell("B2").expect("B2"),
            Some(Cell::Value(Value::Number(number))) if number == &Number::new("42").expect("number")
        ));
        assert!(matches!(
            sheet.cell("C3").expect("C3"),
            Some(Cell::Formula(_))
        ));

        let restored = book
            .apply(&committed.patch().inverse())
            .expect("inverse patch");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored"),
            source_bytes
        );
        assert!(matches!(
            source.apply(committed.patch()),
            Ok(applied) if applied.workbook().sheet("Sheet1").expect("lookup").expect("sheet").cell("A1").expect("cell").is_some()
        ));
        assert!(matches!(
            book.apply(committed.patch()),
            Err(Error::PatchConflict { .. })
        ));
    }

    #[test]
    fn clear_and_remove_have_distinct_missing_and_empty_semantics() {
        let source = Workbook::new().expect("source workbook");
        let mut edit = source.edit().expect("edit");
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .set("A1", "value")
            .expect("set");
        let first = edit.commit().expect("first commit").into_workbook();

        let mut edit = first.edit().expect("edit");
        let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
        sheet.clear("A1").expect("clear");
        sheet.clear("B1").expect("clear missing");
        let cleared = edit.commit().expect("clear commit");
        assert_eq!(cleared.patch().len(), 1);
        let sheet = cleared
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(sheet.cell("A1").expect("cell"), Some(Cell::Empty)));
        assert!(sheet.cell("B1").expect("cell").is_none());

        let mut edit = cleared.workbook().edit().expect("edit");
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .remove("A1")
            .expect("remove");
        let removed = edit.commit().expect("remove commit");
        assert!(
            removed
                .workbook()
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .cell("A1")
                .expect("cell")
                .is_none()
        );
    }

    #[test]
    fn clearing_a_cell_created_in_the_same_transaction_keeps_an_empty_record() {
        let source = Workbook::new().expect("source workbook");
        let mut edit = source.edit().expect("edit");
        let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
        sheet.set("A1", "temporary").expect("set");
        sheet.clear("A1").expect("clear");
        let committed = edit.commit().expect("commit");
        assert!(matches!(
            committed
                .workbook()
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .cell("A1")
                .expect("cell"),
            Some(Cell::Empty)
        ));
    }

    #[test]
    fn independently_prepared_disjoint_edits_join_after_threaded_work() {
        fn assert_send<T: Send>() {}
        assert_send::<Edit>();

        let source = Workbook::new().expect("source workbook");
        let (mut left, right) = std::thread::scope(|scope| {
            let left_source = source.clone();
            let right_source = source.clone();
            let left = scope.spawn(move || {
                let mut edit = left_source.edit().expect("left edit");
                edit.sheet("Sheet1")
                    .expect("lookup")
                    .expect("sheet")
                    .set("A1", "left")
                    .expect("left cell");
                edit
            });
            let right = scope.spawn(move || {
                let mut edit = right_source.edit().expect("right edit");
                edit.sheet(0usize)
                    .expect("lookup")
                    .expect("sheet")
                    .set("C3", 42_i32)
                    .expect("right cell");
                edit
            });
            (
                left.join().expect("left worker"),
                right.join().expect("right worker"),
            )
        });

        left.join(right).expect("disjoint join");
        assert_eq!(left.len(), 2);
        let committed = left.commit().expect("joined commit");
        let sheet = committed
            .workbook()
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(
            sheet.cell("A1").expect("A1"),
            Some(Cell::Value(Value::Text(text))) if text.as_str() == "left"
        ));
        assert!(matches!(
            sheet.cell("C3").expect("C3"),
            Some(Cell::Value(Value::Number(number))) if number.as_str() == "42"
        ));

        let mut empty = source.edit().expect("empty edit");
        let mut incoming = source.edit().expect("incoming edit");
        incoming
            .sheet("Sheet1")
            .expect("lookup")
            .expect("sheet")
            .set("D4", true)
            .expect("incoming cell");
        empty.join(incoming).expect("adopt incoming sheet map");
        assert_eq!(empty.len(), 1);
    }

    #[test]
    fn join_conflicts_are_structured_and_return_the_rejected_edit() {
        let source = Workbook::new().expect("source workbook");
        let mut left = source.edit().expect("left edit");
        let mut left_sheet = left.sheet(0usize).expect("lookup").expect("sheet");
        left_sheet
            .set("C3", "left tail")
            .expect("left tail cell")
            .set("A1", "left")
            .expect("left first cell");
        let mut right = source.edit().expect("right edit");
        let mut right_sheet = right.sheet("Sheet1").expect("lookup").expect("sheet");
        right_sheet
            .set("A1", "right")
            .expect("right first cell")
            .set("C3", "right tail")
            .expect("right tail cell");

        let error = match left.join(right) {
            Ok(_) => panic!("overlapping edits must not join"),
            Err(error) => error,
        };
        assert_eq!(left.len(), 2);
        let conflicts = error.conflicts().expect("overlap details");
        assert_eq!(conflicts.len(), 2);
        assert_eq!(conflicts.conflicts().len(), 1);
        assert_eq!(conflicts.conflicts()[0].sheet(), "Sheet1");
        assert_eq!(conflicts.conflicts()[0].position(), 0);
        assert_eq!(
            conflicts.conflicts()[0].addresses(),
            &[
                Address::from_a1("A1").expect("first address"),
                Address::from_a1("C3").expect("tail address"),
            ]
        );
        let rejected = error.into_rejected();
        assert_eq!(rejected.len(), 2);

        let other_source = Workbook::new().expect("other source");
        let other = other_source.edit().expect("other edit");
        let error = match left.join(other) {
            Ok(_) => panic!("different snapshots must not join"),
            Err(error) => error,
        };
        assert!(matches!(error.failure(), JoinFailure::DifferentSnapshot));
        assert!(error.conflicts().is_none());
        assert_eq!(left.len(), 2);
    }

    #[test]
    fn calculation_chain_removal_is_atomic_and_reversible() {
        let baseline = Workbook::new().expect("baseline");
        let mut package = baseline.inner.package.clone();
        let chain_uri = PackURI::new("/xl/calcChain.xml").expect("chain URI");
        package
            .try_add_part(Box::new(BlobPart::new(
                chain_uri.clone(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
                    .to_owned(),
                br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>"#.to_vec(),
            )))
            .expect("chain part");
        package
            .get_part_mut(&baseline.inner.workbook_uri)
            .expect("workbook part")
            .rels_mut()
            .try_add_relationship(
                litchi_opc::constants::relationship_type::CALC_CHAIN.to_owned(),
                "calcChain.xml".to_owned(),
                "rId3".to_owned(),
                TargetMode::Internal,
            )
            .expect("chain relationship");
        let source = Workbook::from_package(package).expect("workbook with chain");
        let source_bytes = source.to_bytes().expect("source bytes");

        let mut edit = source.edit().expect("edit");
        edit.sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .set("A1", 7_i32)
            .expect("set");
        let committed = edit.commit().expect("commit");
        assert!(
            committed
                .workbook()
                .inner
                .package
                .get_part(&chain_uri)
                .is_err()
        );
        assert!(
            calculation_chain_removal(committed.workbook())
                .expect("chain query")
                .is_empty()
        );

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored"),
            source_bytes
        );
        assert!(
            !calculation_chain_removal(restored.workbook())
                .expect("restored chain")
                .is_empty()
        );

        let mut shared_target = source.inner.package.clone();
        let mut referrer = BlobPart::new(
            PackURI::new("/xl/custom.xml").expect("custom URI"),
            "application/xml".to_owned(),
            b"<custom/>".to_vec(),
        );
        referrer
            .rels_mut()
            .try_add_relationship(
                "urn:litchi:test-reference".to_owned(),
                "calcChain.xml".to_owned(),
                "rId1".to_owned(),
                TargetMode::Internal,
            )
            .expect("extra incoming relationship");
        shared_target
            .try_add_part(Box::new(referrer))
            .expect("referrer part");
        let shared_target = Workbook::from_package(shared_target).expect("shared target workbook");
        assert!(matches!(
            calculation_chain_removal(&shared_target),
            Err(Error::Invalid(message)) if message.contains("another incoming relationship")
        ));
    }

    #[test]
    fn shared_style_crud_is_lineage_checked_reversible_and_exact() {
        let source = styled_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let styles = source.styles().expect("styles");
        assert_eq!(styles.len(), 2);
        let base = styles.base().expect("base style");
        let accent = styles.get(1).expect("accent style");
        assert_eq!(base.fan_out().expect("base fan-out"), 1);
        assert_eq!(accent.fan_out().expect("accent fan-out"), 1);

        let sheet = source.sheet("Sheet1").expect("lookup").expect("sheet");
        assert!(matches!(
            sheet.local_style("B1").expect("local style"),
            Some(crate::LocalStyle::Default)
        ));
        let Some(crate::LocalStyle::Shared(local)) = sheet.local_style("A1").expect("local style")
        else {
            panic!("A1 must have an explicit shared style")
        };
        assert!(local.same(&accent));
        assert!(
            sheet
                .style("B1")
                .expect("resolved style")
                .is_some_and(|style| style.same(&base))
        );

        let mut edit = source.edit().expect("edit");
        let mut sheet = edit.sheet("Sheet1").expect("lookup").expect("sheet");
        sheet
            .set("C1", 42_i32)
            .and_then(|sheet| sheet.style("C1", &accent))
            .and_then(|sheet| sheet.style("D1", &accent))
            .expect("style changes");
        let committed = edit.commit().expect("commit");
        assert_eq!(committed.patch().len(), 2);
        assert!(matches!(
            committed.patch().changes()[0].before(),
            State::Missing
        ));
        assert!(matches!(
            committed.patch().changes()[0].after(),
            State::Cell {
                content: Cell::Value(Value::Number(number)),
                style: StyleState::Shared(_),
            } if number.as_str() == "42"
        ));

        let book = committed.workbook();
        let sheet = book.sheet(0usize).expect("lookup").expect("sheet");
        assert!(matches!(sheet.cell("D1").expect("D1"), Some(Cell::Empty)));
        let styles = book.styles().expect("styles");
        let inherited = styles.find(&accent.key()).expect("inherited style key");
        assert!(inherited.same(&accent));
        assert!(!inherited.same_workbook(&accent));
        assert_eq!(accent.fan_out().expect("source fan-out"), 1);
        assert_eq!(inherited.fan_out().expect("descendant fan-out"), 3);
        assert!(
            sheet
                .style("C1")
                .expect("style")
                .is_some_and(|style| style.same(&inherited))
        );

        let mut descendant = book.edit().expect("descendant edit");
        descendant
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .style("E1", &accent)
            .expect("reuse inherited style lineage");
        let descendant = descendant.commit().expect("descendant commit");
        assert!(matches!(
            descendant
                .workbook()
                .sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .local_style("E1")
                .expect("local style"),
            Some(crate::LocalStyle::Shared(_))
        ));

        let reopened = Workbook::from_bytes(source_bytes.clone()).expect("reopened source");
        assert!(
            reopened
                .styles()
                .expect("reopened styles")
                .find(&accent.key())
                .is_none()
        );
        let replayed = reopened
            .apply(committed.patch())
            .expect("replay onto byte-identical source");
        let State::Cell {
            style: StyleState::Shared(replayed_key),
            ..
        } = replayed.patch().changes()[0].after()
        else {
            panic!("replayed change must retain its shared style")
        };
        assert!(
            replayed
                .workbook()
                .styles()
                .expect("replayed styles")
                .find(replayed_key)
                .is_some()
        );
        assert!(
            book.styles()
                .expect("original lineage")
                .find(replayed_key)
                .is_none()
        );

        let restored = book
            .apply(&committed.patch().inverse())
            .expect("inverse patch");
        assert_eq!(
            restored.workbook().to_bytes().expect("restored bytes"),
            source_bytes
        );

        let foreign = Workbook::new()
            .expect("other workbook")
            .styles()
            .expect("other styles")
            .base()
            .expect("other base style");
        let mut edit = source.edit().expect("edit");
        assert!(matches!(
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .style("A1", &foreign),
            Err(Error::ForeignStyle)
        ));

        let mut changed_package = source.inner.package.clone();
        let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
        let changed_xml = {
            let styles = changed_package.get_part(&styles_uri).expect("styles part");
            std::str::from_utf8(styles.blob())
                .expect("UTF-8 styles")
                .replace("FFFFFF00", "FFFF0000")
                .into_bytes()
        };
        changed_package
            .get_part_mut(&styles_uri)
            .expect("styles part")
            .set_blob(changed_xml);
        let changed = Workbook::from_package_with_styles(changed_package, Some(&source))
            .expect("changed style table");
        assert!(
            changed
                .styles()
                .expect("changed styles")
                .find(&accent.key())
                .is_none()
        );
        let mut edit = changed.edit().expect("changed edit");
        assert!(matches!(
            edit.sheet(0usize)
                .expect("lookup")
                .expect("sheet")
                .style("A1", &accent),
            Err(Error::ForeignStyle)
        ));
        assert!(matches!(
            changed.apply(committed.patch()),
            Err(Error::PatchConflict { part }) if part == "/xl/styles.xml"
        ));
    }

    #[test]
    fn payload_and_style_effects_on_one_cell_join_without_locks() {
        let source = styled_workbook();
        let accent = source.styles().expect("styles").get(1).expect("accent");
        let (mut payload, style) = std::thread::scope(|scope| {
            let payload_source = source.clone();
            let style_source = source.clone();
            let accent = accent.clone();
            let payload = scope.spawn(move || {
                let mut edit = payload_source.edit().expect("payload edit");
                edit.sheet(0usize)
                    .expect("lookup")
                    .expect("sheet")
                    .set("B1", 9_i32)
                    .expect("payload");
                edit
            });
            let style = scope.spawn(move || {
                let mut edit = style_source.edit().expect("style edit");
                edit.sheet("Sheet1")
                    .expect("lookup")
                    .expect("sheet")
                    .style("B1", &accent)
                    .expect("style");
                edit
            });
            (
                payload.join().expect("payload worker"),
                style.join().expect("style worker"),
            )
        });
        payload.join(style).expect("disjoint cell facets");
        assert_eq!(payload.len(), 1);
        let committed = payload.commit().expect("commit");
        assert_eq!(committed.patch().len(), 1);
        assert!(matches!(
            committed.patch().changes()[0].after(),
            State::Cell {
                content: Cell::Value(Value::Number(number)),
                style: StyleState::Shared(_),
            } if number.as_str() == "9"
        ));

        let sheet = committed
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(
            sheet.cell("B1").expect("B1"),
            Some(Cell::Value(Value::Number(number))) if number.as_str() == "9"
        ));
        assert!(matches!(
            sheet.local_style("B1").expect("style"),
            Some(crate::LocalStyle::Shared(_))
        ));
    }

    #[test]
    fn resetting_style_is_distinct_from_removal_and_missing_is_a_no_op() {
        let source = styled_workbook();
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        let mut sheet = edit.sheet(0usize).expect("lookup").expect("sheet");
        sheet
            .reset_style("A1")
            .and_then(|sheet| sheet.reset_style("Z99"))
            .expect("style resets");
        let committed = edit.commit().expect("commit");
        assert_eq!(committed.patch().len(), 1);
        assert!(matches!(
            committed.patch().changes()[0].after(),
            State::Cell {
                style: StyleState::Default,
                ..
            }
        ));
        let sheet = committed
            .workbook()
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet");
        assert!(matches!(
            sheet.local_style("A1").expect("local style"),
            Some(crate::LocalStyle::Default)
        ));
        assert!(sheet.cell("Z99").expect("missing").is_none());

        let restored = committed
            .workbook()
            .apply(&committed.patch().inverse())
            .expect("inverse");
        assert_eq!(restored.workbook().to_bytes().expect("bytes"), source_bytes);
    }

    #[test]
    fn signed_packages_refuse_edits_before_mutation() {
        let baseline = Workbook::new().expect("baseline");
        let mut package = baseline.inner.package.clone();
        package
            .rels_mut()
            .try_add_relationship(
                litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                "_xmlsignatures/origin.sigs".to_owned(),
                "rIdSignature".to_owned(),
                TargetMode::Internal,
            )
            .expect("signature relationship");
        let signed = Workbook::from_package(package).expect("signed workbook snapshot");

        assert!(matches!(signed.edit(), Err(Error::Signed)));
        assert!(matches!(
            signed.apply(&Patch::default()),
            Err(Error::Signed)
        ));
    }

    fn styled_workbook() -> Workbook {
        let baseline = Workbook::new().expect("baseline");
        let mut package = baseline.inner.package.clone();
        package
            .get_part_mut(&PackURI::new("/xl/styles.xml").expect("styles URI"))
            .expect("styles part")
            .set_blob(
                br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#.to_vec(),
            );
        package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
            .expect("worksheet part")
            .set_blob(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#.to_vec(),
            );
        Workbook::from_package(package).expect("styled workbook")
    }
}
