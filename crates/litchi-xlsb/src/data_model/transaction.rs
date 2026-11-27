//! Detached, source-checked edits for XLSB Data Model metadata and payload.

use std::collections::HashMap;
use std::sync::Arc;

use litchi_xldm::identity::{
    Xldm140IdentityProjection, Xldm140RenameError, Xldm140RenameErrorKind,
};
use litchi_xldm::{OlapProofLimits, StorageProfile, prove_xldm140_closure, prove_xldm140_olap};

use super::codec::validate_definition;
use super::model::{Definition, Model, ModelPart, TimeGrouping};
use super::package::{validate_definition_connection_names, validate_payload};
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use crate::package::error::{Error, Result};

crate::part_edit::limits_options!(super::SemanticLimits, "Data Model");

/// A staged Data Model edit of one package. Dropping it discards it.
#[derive(Clone, Debug)]
pub struct Edit {
    source: crate::Package,
    limits: super::SemanticLimits,
    before: Option<Snapshot>,
    definition: Option<Definition>,
    payload: Option<Arc<Vec<u8>>>,
    validated_time_grouping_edit: bool,
    /// Identity state against which subsequent staged metadata edits are
    /// checked.  A successful closure-aware rename advances this baseline;
    /// comparing every later edit with the original snapshot would make a
    /// safe rename followed by an orthogonal metadata edit impossible.
    identity_baseline: Option<Definition>,
    payload_edit_kind: PayloadEditKind,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum PayloadEditKind {
    None,
    OpaqueReplacement,
    IdentityRename,
}

impl Edit {
    pub(crate) fn new(source: crate::Package, limits: super::SemanticLimits) -> Self {
        Self {
            source,
            limits,
            before: None,
            definition: None,
            payload: None,
            validated_time_grouping_edit: false,
            identity_baseline: None,
            payload_edit_kind: PayloadEditKind::None,
        }
    }

    fn ensure_prepared(&mut self) -> Result<()> {
        if self.before.is_some() {
            return Ok(());
        }
        let before = Snapshot::read_with_limits(self.source.opc_package(), self.limits)?;
        self.definition = before.definition().cloned();
        self.payload = before.part().map(|part| Arc::clone(&part.bytes));
        self.identity_baseline = self.definition.clone();
        self.before = Some(before);
        Ok(())
    }

    fn snapshot(&self) -> Result<&Snapshot> {
        self.before
            .as_ref()
            .ok_or_else(|| invalid("Data Model edit is not prepared"))
    }

    /// Immutable source snapshot used by stale-source checks.
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.ensure_prepared()?;
        self.snapshot()
    }

    /// Borrow the currently staged typed workbook metadata.
    pub fn definition(&mut self) -> Result<Option<&Definition>> {
        self.ensure_prepared()?;
        Ok(self.definition.as_ref())
    }

    /// Borrow the currently staged opaque payload bytes.
    pub fn payload(&mut self) -> Result<Option<&[u8]>> {
        self.ensure_prepared()?;
        Ok(self.payload.as_ref().map(|value| value.as_slice()))
    }

    /// Replace typed workbook metadata while retaining the existing payload.
    pub fn set_definition(&mut self, definition: Definition) -> Result<bool> {
        self.ensure_prepared()?;
        if self.payload.is_none() {
            return Err(invalid(
                "cannot set Data Model workbook records without a model payload",
            ));
        }
        validate_definition(&definition, self.limits)?;
        self.validate_connection_closure(&definition)?;
        if self.definition.as_ref() == Some(&definition) {
            return Ok(false);
        }
        ensure_opaque_identity_compatible(self.identity_baseline.as_ref(), &definition)?;
        self.definition = Some(definition);
        Ok(true)
    }

    /// Edit typed workbook metadata through a retry-safe cloned draft.
    pub fn edit_definition(
        &mut self,
        edit: impl FnOnce(&mut Definition) -> Result<()>,
    ) -> Result<bool> {
        self.ensure_prepared()?;
        let mut definition = self.definition.clone().ok_or_else(|| {
            invalid("cannot edit Data Model workbook records when the model is absent")
        })?;
        let before = definition.clone();
        edit(&mut definition)?;
        validate_definition(&definition, self.limits)?;
        self.validate_connection_closure(&definition)?;
        if definition == before {
            return Ok(false);
        }
        ensure_opaque_identity_compatible(self.identity_baseline.as_ref(), &definition)?;
        self.definition = Some(definition);
        Ok(true)
    }

    /// Rename one Data Model table through the complete source-bound XLDM
    /// identity closure.
    ///
    /// The workbook record, every outer relationship endpoint, and any
    /// source-qualified time-grouping reference are staged as one operation.
    /// The opaque model part is rewritten only after the borrowed source has
    /// proved a complete known Xldm140 closure; unknown members, non-Xldm140
    /// payloads, changed TableIDs/column identities, DAX, refresh, and
    /// external-source actions remain outside this API and are refused by the
    /// neutral closure owner.  Errors leave this detached transaction
    /// unchanged.
    pub fn rename_table(&mut self, table_id: &str, new_name: &str) -> Result<bool> {
        self.ensure_prepared()?;
        let current = self
            .definition
            .clone()
            .ok_or_else(|| invalid("cannot rename a Data Model table when the model is absent"))?;
        let table_index = unique_table_index(&current, table_id)?;
        let old_name = current.tables[table_index].name.clone();
        if old_name == new_name {
            return Ok(false);
        }
        if self.payload_edit_kind == PayloadEditKind::OpaqueReplacement {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_data_model",
                "cannot rename a table after an opaque XLDM payload replacement; the replacement has no source identity provenance",
            ));
        }
        validate_table_rename_name(&current, table_index, new_name)?;
        let payload = self
            .payload
            .as_ref()
            .ok_or_else(|| invalid("cannot rename a Data Model table without a model payload"))?;

        // The existing typed grouping records are part of the source identity
        // closure.  Refuse an already-invalid source before the inner writer
        // has a chance to produce a repaired-looking candidate.
        validate_time_groupings(
            payload.as_slice(),
            &current.time_groupings,
            self.limits.max_part_bytes,
        )?;

        // Prove and materialize the inner candidate only after the cheap
        // outer name checks above.  The neutral writer validates the XML name
        // and its bounded path length again before it allocates changed
        // metadata or the rewritten outer storage.
        let renamed_payload = rename_xldm_table_payload(
            payload.as_slice(),
            table_id,
            &current,
            new_name,
            self.limits.max_part_bytes,
        )?;

        // Build and validate the entire outer semantic candidate before
        // mutating any staged field.  The neutral closure has already proved
        // the matching inner XML name and preserved every unsupported byte.
        let replacement_name = new_name.to_owned();
        let mut definition = current;
        definition.tables[table_index].name = replacement_name.clone();
        let old_name_folded = old_name.to_lowercase();
        for relationship in &mut definition.relationships {
            if relationship.from_table.to_lowercase() == old_name_folded {
                relationship.from_table = replacement_name.clone();
            }
            if relationship.to_table.to_lowercase() == old_name_folded {
                relationship.to_table = replacement_name.clone();
            }
        }
        for grouping in &mut definition.time_groupings {
            if grouping.table_name.to_lowercase() == old_name_folded {
                grouping.table_name = replacement_name.clone();
            }
        }
        validate_definition(&definition, self.limits)?;
        self.validate_connection_closure(&definition)?;
        validate_time_groupings(
            &renamed_payload,
            &definition.time_groupings,
            self.limits.max_part_bytes,
        )?;

        self.definition = Some(definition.clone());
        self.identity_baseline = Some(definition);
        self.payload = Some(Arc::new(renamed_payload));
        self.payload_edit_kind = PayloadEditKind::IdentityRename;
        Ok(true)
    }

    /// Add one time grouping after proving its source and every generated
    /// calculated column against the complete source XLDM closure.
    ///
    /// This operation changes only the typed workbook records. The opaque
    /// model payload is retained byte-for-byte; native/generated value
    /// regeneration and calculated-column creation remain a separate writer
    /// operation and are refused when the referenced columns are absent.
    pub fn add_time_grouping(&mut self, grouping: TimeGrouping) -> Result<bool> {
        self.ensure_prepared()?;
        let mut definition = self.current_definition()?;
        if definition.time_groupings.iter().any(|existing| {
            existing.table_name == grouping.table_name && existing.column_id == grouping.column_id
        }) {
            return Err(invalid("duplicate Data Model time grouping"));
        }
        definition.time_groupings.push(grouping);
        self.validate_time_grouping_definition(&definition)?;
        self.definition = Some(definition);
        self.validated_time_grouping_edit = true;
        Ok(true)
    }

    /// Replace one existing time grouping by its qualified source identity.
    ///
    /// The replacement may change calculated-column selections and names, but
    /// every resulting source/calculated identity must already exist in the
    /// source payload. A missing inner column is rejected before the detached
    /// draft is changed.
    pub fn replace_time_grouping(
        &mut self,
        table_name: &str,
        column_id: &str,
        grouping: TimeGrouping,
    ) -> Result<bool> {
        self.ensure_prepared()?;
        let mut definition = self.current_definition()?;
        let Some(index) = definition.time_groupings.iter().position(|existing| {
            existing.table_name == table_name && existing.column_id == column_id
        }) else {
            return Err(invalid("Data Model time grouping. is absent"));
        };
        if definition.time_groupings[index] == grouping {
            return Ok(false);
        }
        if definition
            .time_groupings
            .iter()
            .enumerate()
            .any(|(other, existing)| {
                other != index
                    && existing.table_name == grouping.table_name
                    && existing.column_id == grouping.column_id
            })
        {
            return Err(invalid("duplicate Data Model time grouping"));
        }
        definition.time_groupings[index] = grouping;
        self.validate_time_grouping_definition(&definition)?;
        self.definition = Some(definition);
        self.validated_time_grouping_edit = true;
        Ok(true)
    }

    /// Remove a time grouping by its qualified source identity.
    ///
    /// Removing workbook metadata does not delete the opaque calculated-column
    /// payload. The remaining grouping records are still checked against the
    /// source closure before publication, so a removal cannot conceal an
    /// unrelated dangling reference.
    pub fn remove_time_grouping(&mut self, table_name: &str, column_id: &str) -> Result<bool> {
        self.ensure_prepared()?;
        let mut definition = self.current_definition()?;
        let Some(index) = definition.time_groupings.iter().position(|existing| {
            existing.table_name == table_name && existing.column_id == column_id
        }) else {
            return Ok(false);
        };
        definition.time_groupings.remove(index);
        self.validate_time_grouping_definition(&definition)?;
        self.definition = Some(definition);
        self.validated_time_grouping_edit = true;
        Ok(true)
    }

    fn current_definition(&self) -> Result<Definition> {
        self.definition.clone().ok_or_else(|| {
            invalid("cannot edit Data Model time groupings when the model is absent")
        })
    }

    fn validate_time_grouping_definition(&self, definition: &Definition) -> Result<()> {
        if self.payload.is_none() {
            return Err(invalid(
                "cannot edit Data Model time groupings without a model payload",
            ));
        }
        validate_definition(definition, self.limits)?;
        self.validate_connection_closure(definition)?;
        let Some(payload) = self.payload.as_ref().map(Arc::as_ref) else {
            return Err(invalid(
                "cannot edit Data Model time groupings without a model payload",
            ));
        };
        validate_time_groupings(
            payload,
            &definition.time_groupings,
            self.limits.max_part_bytes,
        )
    }

    /// Replace the inert payload while retaining typed workbook metadata.
    pub fn replace_payload(&mut self, payload: Vec<u8>) -> Result<bool> {
        self.ensure_prepared()?;
        let definition = self
            .definition
            .as_ref()
            .ok_or_else(|| invalid("cannot set a Data Model payload without workbook records"))?;
        let part = ModelPart {
            part_name: super::package::DATA_MODEL_PART_NAME.to_string(),
            content_type: super::package::DATA_MODEL_CONTENT_TYPE.to_string(),
            bytes: Arc::new(payload),
        };
        validate_definition(definition, self.limits)?;
        self.validate_connection_closure(definition)?;
        validate_payload(&part)?;
        if part.bytes.len() > self.limits.max_part_bytes {
            return Err(Error::limit(
                crate::data_model::LimitKind::PayloadBytes,
                part.bytes.len(),
                self.limits.max_part_bytes,
            ));
        }
        if self
            .payload
            .as_ref()
            .is_some_and(|value| value.as_slice() == part.bytes.as_slice())
        {
            return Ok(false);
        }
        if self.identity_baseline.as_ref() != self.snapshot()?.definition() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "edit_data_model",
                "cannot replace an opaque XLDM payload after a staged table-identity rename; use the closure-aware table rename or replace the complete model first",
            ));
        }
        validate_time_groupings(
            part.bytes(),
            &definition.time_groupings,
            self.limits.max_part_bytes,
        )?;
        self.payload = Some(Arc::clone(&part.bytes));
        self.payload_edit_kind = PayloadEditKind::OpaqueReplacement;
        Ok(true)
    }

    /// Replace or remove the complete model pair after validating both owners.
    pub fn replace_model(&mut self, model: Option<Model>) -> Result<bool> {
        self.ensure_prepared()?;
        let current_payload = self.payload.as_ref().map(|payload| payload.as_slice());
        let (definition, payload) = match model {
            Some(model) => {
                validate_definition(&model.definition, self.limits)?;
                self.validate_connection_closure(&model.definition)?;
                validate_payload(&model.part)?;
                if model.part.bytes.len() > self.limits.max_part_bytes {
                    return Err(Error::limit(
                        crate::data_model::LimitKind::PayloadBytes,
                        model.part.bytes.len(),
                        self.limits.max_part_bytes,
                    ));
                }
                validate_time_groupings(
                    model.part.bytes(),
                    &model.definition.time_groupings,
                    self.limits.max_part_bytes,
                )?;
                ensure_opaque_identity_compatible(
                    self.identity_baseline.as_ref(),
                    &model.definition,
                )?;
                if self.identity_baseline.as_ref() != self.snapshot()?.definition()
                    && self.payload.as_ref().map(|payload| payload.as_slice())
                        != Some(model.part.bytes())
                {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "edit_data_model",
                        "cannot replace an opaque XLDM payload after a staged table-identity rename; use the closure-aware table rename or replace the complete model first",
                    ));
                }
                (Some(model.definition), Some(Arc::clone(&model.part.bytes)))
            },
            None => (None, None),
        };
        if self.definition == definition && self.payload == payload {
            return Ok(false);
        }
        let payload_changed = current_payload != payload.as_ref().map(|value| value.as_slice());
        self.definition = definition;
        self.payload = payload;
        if payload_changed {
            self.payload_edit_kind = PayloadEditKind::OpaqueReplacement;
        }
        Ok(true)
    }

    /// Remove the complete model pair.
    pub fn remove_model(&mut self) -> Result<bool> {
        self.replace_model(None)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        self.ensure_prepared()?;
        let definition = self.snapshot()?.definition().cloned();
        let payload = self.snapshot()?.part().map(|part| Arc::clone(&part.bytes));
        self.definition = definition;
        self.payload = payload;
        self.identity_baseline = self.definition.clone();
        self.validated_time_grouping_edit = false;
        self.payload_edit_kind = PayloadEditKind::None;
        Ok(())
    }

    /// Validate the staged Data Model, publish it to a new package and read
    /// it back. An unchanged model returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed edit of a signed workbook, the
    /// closure, identity and payload refusals, and the workbook validation
    /// errors of the result.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_prepared()?;
        let before = self.snapshot()?;
        let after_model = match (&self.definition, &self.payload) {
            (None, None) => None,
            (Some(definition), Some(payload)) => {
                let (part_name, content_type) = before
                    .part()
                    .map(|part| (part.part_name.clone(), part.content_type.clone()))
                    .unwrap_or_else(|| {
                        (
                            super::package::DATA_MODEL_PART_NAME.to_string(),
                            super::package::DATA_MODEL_CONTENT_TYPE.to_string(),
                        )
                    });
                Some(Model {
                    definition: definition.clone(),
                    part: ModelPart {
                        part_name,
                        content_type,
                        bytes: Arc::clone(payload),
                    },
                })
            },
            (Some(_), None) => {
                return Err(invalid(
                    "Data Model workbook records are present without a payload",
                ));
            },
            (None, Some(_)) => {
                return Err(invalid(
                    "Data Model payload is present without workbook records",
                ));
            },
        };
        let unchanged = match (before.model(), after_model.as_ref()) {
            (None, None) => true,
            (Some(before), Some(after)) => before == after,
            _ => false,
        };
        if unchanged {
            let before = self
                .before
                .take()
                .ok_or_else(|| invalid("Data Model edit is not prepared"))?;
            return Ok(Commit::new(
                self.source,
                Patch::new(before.clone(), before),
                litchi_core::edit::Diagnostics::UNCHANGED,
            ));
        }
        let after_definition = after_model.as_ref().map(|model| &model.definition);
        if let Some(after_definition) = after_definition {
            self.validate_connection_closure(after_definition)?;
            if self.validated_time_grouping_edit {
                let before_payload = before.part().map(ModelPart::bytes);
                let after_payload = self.payload.as_ref().map(|value| value.as_slice());
                if self.payload_edit_kind != PayloadEditKind::IdentityRename
                    && before_payload != after_payload
                {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedContent,
                        "edit_data_model",
                        "time-grouping edits cannot be combined with an opaque XLDM payload replacement; native/generated regeneration requires a closure-aware writer",
                    ));
                }
                ensure_opaque_identity_compatible_except_time_groupings(
                    self.identity_baseline.as_ref(),
                    after_definition,
                )?;
                self.validate_time_grouping_definition(after_definition)?;
            } else {
                ensure_opaque_identity_compatible(
                    self.identity_baseline.as_ref(),
                    after_definition,
                )?;
            }
            if let Some(after_payload) = self.payload.as_ref() {
                validate_time_groupings(
                    after_payload,
                    &after_definition.time_groupings,
                    self.limits.max_part_bytes,
                )?;
            }
        }
        if before.package().is_signed() || before.package().requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let mut candidate = before.package().as_ref().clone();
        super::patch::materialize_models(&mut candidate, before, after_model.as_ref())?;
        let after = Snapshot::read_with_limits(&candidate, before.limits())?;
        if after.model() != after_model.as_ref() {
            return Err(invalid(
                "Data Model transaction readback did not match staged model",
            ));
        }
        if after.same_source(before) {
            return Err(invalid(
                "changed Data Model transaction produced no source change",
            ));
        }
        let package = self.source.reparse_candidate(candidate)?;
        let before = self
            .before
            .take()
            .ok_or_else(|| invalid("Data Model edit is not prepared"))?;
        Ok(Commit::new(
            package,
            Patch::new(before, after),
            litchi_core::edit::Diagnostics::CHANGED,
        ))
    }

    fn validate_connection_closure(&self, definition: &Definition) -> Result<()> {
        validate_definition_connection_names(
            definition,
            self.snapshot()?
                .connection_names()
                .map(|names| names.iter().map(String::as_str)),
        )
    }
}

fn validate_time_groupings(
    payload: &[u8],
    groupings: &[TimeGrouping],
    max_part_bytes: usize,
) -> Result<()> {
    for grouping in groupings {
        super::proof::prove_time_grouping(payload, grouping, max_part_bytes)?;
    }
    Ok(())
}

fn unique_table_index(definition: &Definition, table_id: &str) -> Result<usize> {
    let mut matches = definition
        .tables
        .iter()
        .enumerate()
        .filter(|(_, table)| table.id == table_id)
        .map(|(index, _)| index);
    let Some(index) = matches.next() else {
        return Err(invalid("Data Model table identity is absent"));
    };
    if matches.next().is_some() {
        return Err(invalid("Data Model table identity is ambiguous"));
    }
    Ok(index)
}

fn rename_xldm_table_payload(
    payload: &[u8],
    table_id: &str,
    outer: &Definition,
    new_name: &str,
    max_part_bytes: usize,
) -> Result<Vec<u8>> {
    let limits = litchi_xldm::Limits::new(max_part_bytes, max_part_bytes);
    let storage =
        litchi_xldm::inspect_shared(payload, &limits).map_err(super::proof::map_xldm_error)?;
    if storage.profile() != StorageProfile::Xldm140 {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "Data Model table rename requires a complete Xldm140 payload",
        ));
    }
    let metadata = litchi_xldm::metadata::inspect(&storage, &limits)
        .map_err(super::proof::map_xldm_metadata_error)?;
    let native = litchi_xldm::native::inspect(&storage, &metadata.native_parse_options())
        .map_err(|_error| invalid("XLDM native identity proof failed"))?;
    let generated = litchi_xldm::generated::inspect_system_generated(&storage)
        .map_err(|_error| invalid("XLDM generated identity proof failed"))?;
    let olap = litchi_xldm::olap::inspect(&storage, &metadata)
        .map_err(|_error| invalid("XLDM OLAP identity proof failed"))?;
    let olap_proof = prove_xldm140_olap(&storage, &metadata, &olap, OlapProofLimits::default())
        .map_err(super::proof::map_xldm_olap_proof_error)?;
    if !olap_proof.is_complete() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "Data Model table rename requires a complete XLDM OLAP closure",
        ));
    }
    let closure = prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
        .map_err(|_error| invalid("XLDM identity closure proof failed"))?;
    if !closure.is_complete() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "Data Model table rename requires a complete known XLDM identity closure",
        ));
    }
    validate_outer_identity_graph(outer, closure.projection())?;
    if closure.projection().table(table_id).is_none() {
        return Err(invalid("XLDM table identity is absent from the closure"));
    }
    let patch = closure
        .rename_table_name_with_relationships_with_limit(table_id, new_name, max_part_bytes)
        .map_err(|error| map_xldm_rename_error("XLDM table rename", error))?;
    if patch.is_noop() {
        return Err(invalid(
            "XLDM table rename unexpectedly produced an exact no-op",
        ));
    }
    if patch.after().len() > max_part_bytes {
        return Err(Error::limit(
            crate::data_model::LimitKind::PayloadBytes,
            patch.after().len(),
            max_part_bytes,
        ));
    }
    Ok(patch.after().to_vec())
}

/// Check the complete outer workbook identity against the borrowed XLDM
/// projection before the neutral writer is allowed to rewrite any member.
/// Matching by only the selected table would permit a missing or extra outer
/// relationship to survive an otherwise valid inner rename.
fn validate_outer_identity_graph(
    definition: &Definition,
    projection: &Xldm140IdentityProjection,
) -> Result<()> {
    if definition.tables.len() != projection.tables.len() {
        return Err(invalid(
            "outer Data Model table count does not match the proven XLDM identity graph",
        ));
    }
    if definition.relationships.len() != projection.relationships.len() {
        return Err(invalid(
            "outer Data Model relationship count does not match the proven XLDM identity graph",
        ));
    }

    let mut outer_tables = HashMap::<&str, &str>::new();
    outer_tables
        .try_reserve(definition.tables.len())
        .map_err(|source| Error::Allocation {
            resource: "Data Model outer identity tables",
            source,
        })?;
    for table in &definition.tables {
        if outer_tables
            .insert(table.id.as_str(), table.name.as_str())
            .is_some()
        {
            return Err(invalid("outer Data Model table identity is duplicated"));
        }
    }

    let mut proven_table_names = HashMap::<&str, &str>::new();
    proven_table_names
        .try_reserve(projection.tables.len())
        .map_err(|source| Error::Allocation {
            resource: "Data Model proven identity tables",
            source,
        })?;
    for table in &projection.tables {
        let Some(outer_name) = outer_tables.get(table.table_id.as_str()) else {
            return Err(invalid("proven XLDM table identity has no outer table"));
        };
        if outer_name.to_lowercase() != table.xml_name.to_lowercase() {
            return Err(invalid(
                "outer table name does not match the proven XLDM name",
            ));
        }
        if proven_table_names
            .insert(table.table_id.as_str(), table.xml_name.as_str())
            .is_some()
        {
            return Err(invalid("proven XLDM table identity is duplicated"));
        }
    }
    if proven_table_names.len() != outer_tables.len() {
        return Err(invalid(
            "outer Data Model table identity graph contains an extra table",
        ));
    }

    // Relationship identities have no workbook-side ID.  Their complete
    // endpoint tuple is the stable outer identity; compare a case-folded
    // multiset so source lexical casing does not create a false mismatch.
    let mut proven_relationships = HashMap::<(String, String, String, String), usize>::new();
    proven_relationships
        .try_reserve(projection.relationships.len())
        .map_err(|source| Error::Allocation {
            resource: "Data Model proven identity relationships",
            source,
        })?;
    for relationship in &projection.relationships {
        let Some(containing_table_name) =
            proven_table_names.get(relationship.containing_table.as_str())
        else {
            return Err(invalid(
                "proven XLDM relationship has no containing table identity",
            ));
        };
        let key = (
            containing_table_name.to_lowercase(),
            relationship.foreign_column.to_lowercase(),
            relationship.primary_table.to_lowercase(),
            relationship.primary_column.to_lowercase(),
        );
        let count = proven_relationships.entry(key).or_insert(0);
        *count = count
            .checked_add(1)
            .ok_or_else(|| invalid("proven XLDM relationship identity count overflow"))?;
    }
    for relationship in &definition.relationships {
        let key = (
            relationship.from_table.to_lowercase(),
            relationship.from_column.to_lowercase(),
            relationship.to_table.to_lowercase(),
            relationship.to_column.to_lowercase(),
        );
        let Some(count) = proven_relationships.get_mut(&key) else {
            return Err(invalid(
                "outer Data Model relationship is absent from the proven XLDM identity graph",
            ));
        };
        if *count == 0 {
            return Err(invalid(
                "outer Data Model relationship occurs more often than in the proven XLDM identity graph",
            ));
        }
        *count -= 1;
    }
    if proven_relationships.values().any(|count| *count != 0) {
        return Err(invalid(
            "proven XLDM relationship identity graph contains an extra relationship",
        ));
    }
    Ok(())
}

// Keep this admission bound equal to litchi-xldm's private MAX_PATH_BYTES.
// The neutral owner repeats the check, but rejecting before inspect_shared
// avoids allocating a closure for an obviously impossible candidate.
const MAX_XLDM_TABLE_NAME_BYTES: usize = 32 * 1024;

fn validate_table_rename_name(
    definition: &Definition,
    table_index: usize,
    new_name: &str,
) -> Result<()> {
    if new_name.is_empty() {
        return Err(invalid("Data Model table name must not be empty"));
    }
    if !new_name.chars().all(valid_xml10_char) {
        return Err(invalid(
            "Data Model table name contains an XML 1.0-forbidden character",
        ));
    }
    if new_name.len() > MAX_XLDM_TABLE_NAME_BYTES {
        return Err(Error::limit(
            crate::data_model::LimitKind::TableNameBytes,
            new_name.len(),
            MAX_XLDM_TABLE_NAME_BYTES,
        ));
    }
    let folded = new_name.to_lowercase();
    if definition
        .tables
        .iter()
        .enumerate()
        .any(|(index, table)| index != table_index && table.name.to_lowercase() == folded)
    {
        return Err(invalid(
            "Data Model table rename would collide with another outer table name",
        ));
    }
    Ok(())
}

fn valid_xml10_char(character: char) -> bool {
    matches!(
        character,
        '\u{0009}' | '\u{000A}' | '\u{000D}'
            | '\u{0020}'..='\u{D7FF}'
            | '\u{E000}'..='\u{FFFD}'
            | '\u{10000}'..='\u{10FFFF}'
    )
}

fn invalid(message: &'static str) -> Error {
    Error::InvalidFormat(message)
}

fn map_xldm_rename_error(_context: &str, error: Xldm140RenameError) -> Error {
    match error.kind() {
        Xldm140RenameErrorKind::InvalidSource => {
            invalid("the Data Model rename source is not valid")
        },
        Xldm140RenameErrorKind::Unsupported => Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "the feature is not supported",
        ),
        Xldm140RenameErrorKind::LimitExceeded => {
            let Some((actual, maximum)) = error.limit_bounds() else {
                return invalid("the Data Model rename source is not valid");
            };
            match error.limit_kind() {
                Some(kind) => Error::limit(super::LimitKind::Model(kind), actual, maximum),
                None => invalid("the Data Model rename limit has no dimension"),
            }
        },
        Xldm140RenameErrorKind::Allocation => {
            let Some(source) = error.allocation_source() else {
                // The legacy OLAP proof adapter carries only a detail string.
                // Do not synthesize a TryReserveError; retain the diagnostic
                // and fail as an invalid/unproven source instead.
                return invalid("; neutral allocation resource has no allocator source");
            };
            Error::Allocation {
                resource: "Data Model XLDM rename allocation",
                source: source.clone(),
            }
        },
        _ => invalid("the Data Model rename source is not valid"),
    }
}

fn ensure_opaque_identity_compatible(
    before: Option<&Definition>,
    after: &Definition,
) -> Result<()> {
    let Some(before) = before else {
        return Ok(());
    };
    if before.tables != after.tables
        || before.relationships != after.relationships
        || before.time_groupings != after.time_groupings
    {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "cannot rename or restructure Data Model metadata while the MS-XLDM payload is opaque; replace and validate the complete model in a format-aware owner first",
        ));
    }
    Ok(())
}

fn ensure_opaque_identity_compatible_except_time_groupings(
    before: Option<&Definition>,
    after: &Definition,
) -> Result<()> {
    let Some(before) = before else {
        return Ok(());
    };
    if before.tables != after.tables || before.relationships != after.relationships {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "time-grouping edits cannot change Data Model tables or relationships while the MS-XLDM payload is opaque",
        ));
    }
    Ok(())
}

#[cfg(test)]
mod integration_limit_tests {
    use super::*;

    #[test]
    fn olap_proof_dimensions_and_observations_survive_the_package_adapter() {
        use litchi_xldm::LimitKind;
        for kind in [
            LimitKind::Items,
            LimitKind::StringBytes,
            LimitKind::SourceBytes,
            LimitKind::Work,
            LimitKind::OutputBytes,
            LimitKind::PathTextBytes,
        ] {
            let error = crate::data_model::proof::map_xldm_olap_proof_error(
                litchi_xldm::OlapProofError::Limit {
                    kind,
                    observed: 17,
                    maximum: 16,
                },
            );
            assert!(matches!(error, Error::Limit {
                kind: crate::package::LimitKind::DataModel(super::super::LimitKind::Model(actual)),
                observed: 17, maximum: 16,
            } if actual == kind));
        }
    }
}
