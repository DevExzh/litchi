//! Isolated cell transactions, atomic commits, and reversible exact patches.

use std::collections::BTreeMap;
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, Relationship};
use litchi_sheet::{At, Cell as Address};

use super::{Sheet, SheetKind, SheetSelector, Workbook};
use crate::cell::{Cell, Content};
use crate::error::{EditBlock, Error, Result, invalid};
use crate::raw;
use crate::raw::worksheet::edit::Action;

/// Cell state recorded before or after one semantic change.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum State {
    Missing,
    Cell(Cell),
}

impl State {
    fn read(value: Option<&Cell>) -> Self {
        value.map_or(Self::Missing, |cell| Self::Cell(cell.clone()))
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

#[derive(Debug, Clone)]
struct PartChange {
    uri: PackURI,
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
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
        }
    }

    pub(super) fn apply_to(&self, workbook: &Workbook) -> Result<Commit> {
        ensure_unsigned(workbook)?;
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
        let workbook = Workbook::from_package(package)?;
        Ok(Commit {
            workbook,
            patch: self.clone(),
        })
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

    /// Validate and atomically publish a new immutable snapshot.
    pub fn commit(self) -> Result<Commit> {
        ensure_unsigned(&self.base)?;
        if self.is_empty() {
            return Ok(Commit {
                workbook: self.base,
                patch: Patch::default(),
            });
        }
        let mut changes = Vec::new();
        let mut parts = Vec::new();
        let graph = calculation_chain_removal(&self.base)?;
        for (position, requested) in &self.sheets {
            let data = self
                .base
                .inner
                .sheets
                .get(*position)
                .cloned()
                .ok_or_else(|| invalid(format!("edited sheet position {position} disappeared")))?;
            if data.kind != SheetKind::Worksheet {
                return Err(Error::NotWorksheet {
                    sheet: data.name.clone(),
                });
            }
            let sheet = Sheet {
                owner: Arc::clone(&self.base.inner),
                data: Arc::clone(&data),
            };
            let store = sheet.store()?;
            let mut effective = BTreeMap::new();
            for (address, action) in requested {
                let before = store.get(*address);
                if before.is_some_and(|cell| matches!(cell, Cell::Unknown(_)))
                    && !matches!(action, Action::Remove)
                {
                    return Err(Error::EditBlocked {
                        sheet: data.name.clone(),
                        address: *address,
                        reason: EditBlock::UnknownCell,
                    });
                }
                let before_state = State::read(before);
                let after_state = match action {
                    Action::Set(content) => State::Cell(content.as_cell()),
                    Action::Clear => State::Cell(Cell::Empty),
                    Action::ClearIfPresent if before.is_some() => State::Cell(Cell::Empty),
                    Action::ClearIfPresent | Action::Remove => State::Missing,
                };
                if before_state == after_state {
                    continue;
                }
                effective.insert(*address, action.clone());
                changes.push(Change {
                    sheet: data.name.clone().into_boxed_str(),
                    address: *address,
                    before: before_state,
                    after: after_state,
                });
            }
            if effective.is_empty() {
                continue;
            }

            let part = self.base.inner.package.get_part(&data.part_uri)?;
            let before = part.blob_arc();
            let after = raw::worksheet::edit::rewrite(&before, &data.name, &effective)?;
            let parsed = raw::worksheet::parse(&after, || self.base.inner.shared_strings())?;
            for change in changes.iter().rev().take(effective.len()) {
                let actual = State::read(parsed.get(change.address));
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
                workbook: self.base,
                patch: Patch::default(),
            });
        }
        let workbook_part = self
            .base
            .inner
            .package
            .get_part(&self.base.inner.workbook_uri)?;
        let before = workbook_part.blob_arc();
        let after = raw::recalc::invalidate(&before)?;
        if after.as_slice() != before.as_slice() {
            parts.push(PartChange {
                uri: self.base.inner.workbook_uri.clone(),
                before,
                after: Arc::new(after),
            });
        }
        let mut package = self.base.inner.package.clone();
        for part in &parts {
            package
                .get_part_mut(&part.uri)?
                .set_blob_shared(Arc::clone(&part.after));
        }
        for change in &graph {
            change.validate(&package)?;
            change.apply(&mut package)?;
        }
        let workbook = Workbook::from_package(package)?;
        Ok(Commit {
            workbook,
            patch: Patch {
                changes: changes.into_boxed_slice(),
                parts: parts.into_boxed_slice(),
                graph: graph.into_boxed_slice(),
            },
        })
    }

    fn actions(&mut self, position: usize) -> &mut BTreeMap<Address, Action> {
        self.sheets.entry(position).or_default()
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
        self.edit
            .actions(self.position)
            .insert(address, Action::Set(content));
        Ok(self)
    }

    /// Remove primary value/formula content while retaining the cell record,
    /// local style, metadata, comments, and unknown non-payload children.
    pub fn clear<'a>(&mut self, at: impl Into<At<'a>>) -> Result<&mut Self> {
        let address = at.into().resolve()?;
        let actions = self.edit.actions(self.position);
        let action = if matches!(actions.get(&address), Some(Action::Set(_) | Action::Clear)) {
            Action::Clear
        } else {
            Action::ClearIfPresent
        };
        actions.insert(address, action);
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
                "rId2".to_owned(),
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
}
