//! Detached, source-checked edits for XLSB Data Model metadata and payload.

use std::sync::Arc;

use super::codec::validate_definition;
use super::model::{Definition, Model, ModelPart, TimeGrouping};
use super::package::{validate_definition_connection_names, validate_payload};
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use crate::package::error::{Error, Result};

/// A bounded draft detached from an immutable Data Model snapshot.
#[derive(Clone, Debug)]
pub struct Transaction {
    before: Snapshot,
    definition: Option<Definition>,
    payload: Option<Arc<Vec<u8>>>,
    validated_time_grouping_edit: bool,
}

impl Transaction {
    pub(crate) fn new(before: Snapshot) -> Self {
        let definition = before.definition().cloned();
        let payload = before.part().map(|part| Arc::clone(&part.bytes));
        Self {
            before,
            definition,
            payload,
            validated_time_grouping_edit: false,
        }
    }

    /// Immutable source snapshot used by stale-source checks.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged typed workbook metadata.
    pub fn definition(&self) -> Option<&Definition> {
        self.definition.as_ref()
    }

    /// Borrow the currently staged opaque payload bytes.
    pub fn payload(&self) -> Option<&[u8]> {
        self.payload.as_ref().map(|value| value.as_slice())
    }

    /// Replace typed workbook metadata while retaining the existing payload.
    pub fn set_definition(&mut self, definition: Definition) -> Result<bool> {
        if self.payload.is_none() {
            return Err(invalid(
                "cannot set Data Model workbook records without a model payload",
            ));
        }
        validate_definition(&definition, self.before.limits())?;
        self.validate_connection_closure(&definition)?;
        if self.definition.as_ref() == Some(&definition) {
            return Ok(false);
        }
        ensure_opaque_identity_compatible(self.before.definition(), &definition)?;
        self.definition = Some(definition);
        Ok(true)
    }

    /// Edit typed workbook metadata through a retry-safe cloned draft.
    pub fn edit_definition(
        &mut self,
        edit: impl FnOnce(&mut Definition) -> Result<()>,
    ) -> Result<bool> {
        let mut definition = self.definition.clone().ok_or_else(|| {
            invalid("cannot edit Data Model workbook records when the model is absent")
        })?;
        let before = definition.clone();
        edit(&mut definition)?;
        validate_definition(&definition, self.before.limits())?;
        self.validate_connection_closure(&definition)?;
        if definition == before {
            return Ok(false);
        }
        ensure_opaque_identity_compatible(self.before.definition(), &definition)?;
        self.definition = Some(definition);
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
        let mut definition = self.current_definition()?;
        let Some(index) = definition.time_groupings.iter().position(|existing| {
            existing.table_name == table_name && existing.column_id == column_id
        }) else {
            return Err(invalid(format!(
                "Data Model time grouping {table_name}.{column_id} is absent"
            )));
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
        validate_definition(definition, self.before.limits())?;
        self.validate_connection_closure(definition)?;
        let Some(payload) = self.payload.as_ref().map(Arc::as_ref) else {
            return Err(invalid(
                "cannot edit Data Model time groupings without a model payload",
            ));
        };
        validate_time_groupings(payload, &definition.time_groupings)
    }

    /// Replace the inert payload while retaining typed workbook metadata.
    pub fn replace_payload(&mut self, payload: Vec<u8>) -> Result<bool> {
        let definition = self
            .definition
            .as_ref()
            .ok_or_else(|| invalid("cannot set a Data Model payload without workbook records"))?;
        let part = ModelPart {
            part_name: super::package::DATA_MODEL_PART_NAME.to_string(),
            content_type: super::package::DATA_MODEL_CONTENT_TYPE.to_string(),
            bytes: Arc::new(payload),
        };
        validate_definition(definition, self.before.limits())?;
        self.validate_connection_closure(definition)?;
        validate_payload(&part)?;
        if part.bytes.len() > self.before.limits().max_part_bytes {
            return Err(Error::LimitExceeded {
                resource: "Data Model payload bytes",
                actual: part.bytes.len(),
                maximum: self.before.limits().max_part_bytes,
            });
        }
        if self
            .payload
            .as_ref()
            .is_some_and(|value| value.as_slice() == part.bytes.as_slice())
        {
            return Ok(false);
        }
        validate_time_groupings(part.bytes(), &definition.time_groupings)?;
        self.payload = Some(Arc::clone(&part.bytes));
        Ok(true)
    }

    /// Replace or remove the complete model pair after validating both owners.
    pub fn replace_model(&mut self, model: Option<Model>) -> Result<bool> {
        let (definition, payload) = match model {
            Some(model) => {
                validate_definition(&model.definition, self.before.limits())?;
                self.validate_connection_closure(&model.definition)?;
                validate_payload(&model.part)?;
                if model.part.bytes.len() > self.before.limits().max_part_bytes {
                    return Err(Error::LimitExceeded {
                        resource: "Data Model payload bytes",
                        actual: model.part.bytes.len(),
                        maximum: self.before.limits().max_part_bytes,
                    });
                }
                validate_time_groupings(model.part.bytes(), &model.definition.time_groupings)?;
                if let Some(before) = self.before.definition() {
                    ensure_opaque_identity_compatible(Some(before), &model.definition)?;
                }
                (Some(model.definition), Some(Arc::clone(&model.part.bytes)))
            },
            None => (None, None),
        };
        if self.definition == definition && self.payload == payload {
            return Ok(false);
        }
        self.definition = definition;
        self.payload = payload;
        Ok(true)
    }

    /// Remove the complete model pair.
    pub fn remove_model(&mut self) -> Result<bool> {
        self.replace_model(None)
    }

    /// Commit the detached draft into a reversible source-checked patch.
    pub fn commit(self) -> Result<Commit> {
        let before_definition = self.before.definition();
        let after_model = match (&self.definition, &self.payload) {
            (None, None) => None,
            (Some(definition), Some(payload)) => {
                let (part_name, content_type) = self
                    .before
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
        let unchanged = match (self.before.model(), after_model.as_ref()) {
            (None, None) => true,
            (Some(before), Some(after)) => before == after,
            _ => false,
        };
        if unchanged {
            let before = self.before;
            return Ok(Commit::new(Patch::new(before.clone(), before), false));
        }
        let after_definition = after_model.as_ref().map(|model| &model.definition);
        if let Some(after_definition) = after_definition {
            self.validate_connection_closure(after_definition)?;
            if self.validated_time_grouping_edit {
                let before_payload = self.before.part().map(ModelPart::bytes);
                let after_payload = self.payload.as_ref().map(|value| value.as_slice());
                if before_payload != after_payload {
                    return Err(Error::UnsupportedFeature(
                        "time-grouping edits cannot be combined with an opaque XLDM payload replacement; native/generated regeneration requires a closure-aware writer".to_string(),
                    ));
                }
                ensure_opaque_identity_compatible_except_time_groupings(
                    before_definition,
                    after_definition,
                )?;
                self.validate_time_grouping_definition(after_definition)?;
            } else {
                ensure_opaque_identity_compatible(before_definition, after_definition)?;
            }
            if let Some(after_payload) = self.payload.as_ref() {
                validate_time_groupings(after_payload, &after_definition.time_groupings)?;
            }
        }
        if self.before.package().is_signed()
            || self.before.package().requires_signature_edit_policy()
        {
            return Err(Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }
        let mut candidate = self.before.package().as_ref().clone();
        super::patch::materialize_models(&mut candidate, &self.before, after_model.as_ref())?;
        let after = Snapshot::read_with_limits(&candidate, self.before.limits())?;
        if after.model() != after_model.as_ref() {
            return Err(invalid(
                "Data Model transaction readback did not match staged model",
            ));
        }
        if after.same_source(&self.before) {
            return Err(invalid(
                "changed Data Model transaction produced no source change",
            ));
        }
        Ok(Commit::new(Patch::new(self.before, after), true))
    }

    fn validate_connection_closure(&self, definition: &Definition) -> Result<()> {
        validate_definition_connection_names(
            definition,
            self.before
                .connection_names()
                .map(|names| names.iter().map(String::as_str)),
        )
    }
}

fn validate_time_groupings(payload: &[u8], groupings: &[TimeGrouping]) -> Result<()> {
    for grouping in groupings {
        super::proof::prove_time_grouping(payload, grouping)?;
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
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
        return Err(Error::UnsupportedFeature(
            "cannot rename or restructure Data Model metadata while the MS-XLDM payload is opaque; replace and validate the complete model in a format-aware owner first".to_string(),
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
        return Err(Error::UnsupportedFeature(
            "time-grouping edits cannot change Data Model tables or relationships while the MS-XLDM payload is opaque".to_string(),
        ));
    }
    Ok(())
}
