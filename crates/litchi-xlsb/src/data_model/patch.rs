//! Reversible, source-checked Data Model publication.

use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI};

use super::codec;
use super::package::{DATA_MODEL_CONTENT_TYPE, reject_inbound_relationships};
use super::snapshot::Snapshot;
use crate::package::error::{Error, Result};

/// A reversible replacement of the workbook Data Model owner pair.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source state required before publication.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by publication.
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether publication changes no workbook or model-part source bytes.
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return a patch that restores the source state represented by `before`.
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    pub(crate) fn limits(&self) -> super::ReadLimits {
        self.before.limits()
    }

    /// Check the current package against this patch's exact owned source.
    pub(crate) fn check_source(&self, package: &OpcPackage) -> Result<Snapshot> {
        let current = Snapshot::read_with_limits(package, self.limits())?;
        if !current.same_source(&self.before) {
            return Err(invalid("Data Model patch source is stale"));
        }
        Ok(current)
    }

    /// Apply this patch directly to an OPC package after an exact source check.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        let current = self.check_source(package)?;
        if self.is_empty() {
            return Ok(current);
        }
        if package.is_signed() || package.requires_signature_edit_policy() {
            return Err(Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }
        let mut candidate = package.clone();
        self.materialize(&mut candidate)?;
        let resulting = Snapshot::read_with_limits(&candidate, self.limits())?;
        if &resulting != self.after() {
            return Err(invalid(
                "Data Model publication changed the planned semantic or owned graph",
            ));
        }
        *package = candidate;
        Ok(resulting)
    }

    pub(crate) fn materialize(&self, package: &mut OpcPackage) -> Result<()> {
        if self.is_empty() {
            return Ok(());
        }
        materialize_models(package, &self.before, self.after.model())?;
        restore_source_tokens(package, self.after.source(), self.limits())
    }
}

pub(crate) fn materialize_models(
    package: &mut OpcPackage,
    before: &Snapshot,
    after_model: Option<&super::model::Model>,
) -> Result<()> {
    let workbook_name = {
        let workbook = package.main_document_part()?;
        workbook.partname().clone()
    };
    let expected_workbook = PackURI::new(before.source().workbook_name())
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
    if !workbook_name.is_equivalent_to(&expected_workbook) {
        return Err(invalid("Data Model workbook owner changed"));
    }
    let before_definition = before.definition();
    let after_definition = after_model.map(|model| &model.definition);
    if before_definition != after_definition {
        let workbook_bytes = package.get_part(&workbook_name)?.blob().to_vec();
        let workbook_bytes = codec::patch_workbook(
            &workbook_bytes,
            before_definition,
            after_definition,
            before.limits(),
        )?;
        package
            .get_part_mut(&workbook_name)?
            .set_blob(workbook_bytes);
    }

    match after_model.map(|model| &model.part) {
        Some(after_part) => {
            let current_name = before
                .source()
                .model_part()
                .map(|part| PackURI::new(&part.name))
                .transpose()
                .map_err(|error| Error::InvalidUri(error.to_string()))?;
            if let Some(current_name) = current_name {
                let current = package.get_part(&current_name)?;
                let physical_name = current.partname().clone();
                if current.content_type() != DATA_MODEL_CONTENT_TYPE {
                    return Err(invalid(
                        "fixed Data Model part name has a different content type",
                    ));
                }
                if !current.rels().is_empty() {
                    return Err(invalid(
                        "fixed Data Model part has forbidden outbound relationships",
                    ));
                }
                reject_inbound_relationships(package, &physical_name)?;
                let part = package.get_part_mut(&physical_name)?;
                part.set_blob_shared(Arc::clone(&after_part.bytes));
                part.set_content_type(after_part.content_type.clone())?;
            } else {
                let uri = PackURI::new(&after_part.part_name)
                    .map_err(|error| Error::InvalidUri(error.to_string()))?;
                reject_inbound_relationships(package, &uri)?;
                let part = BlobPart::new_shared(
                    uri.clone(),
                    after_part.content_type.clone(),
                    Arc::clone(&after_part.bytes),
                );
                package.try_add_part(Box::new(part))?;
            }
        },
        None => {
            if let Some(before_part) = before.source().model_part() {
                let requested = PackURI::new(&before_part.name)
                    .map_err(|error| Error::InvalidUri(error.to_string()))?;
                let physical_name = package.get_part(&requested)?.partname().clone();
                reject_inbound_relationships(package, &physical_name)?;
                if !package.remove_part(&physical_name) {
                    return Err(invalid("fixed Data Model part disappeared during removal"));
                }
            }
        },
    }
    Ok(())
}

fn restore_source_tokens(
    package: &mut OpcPackage,
    source: &super::snapshot::SourceState,
    limits: super::ReadLimits,
) -> Result<()> {
    let opc_limits = super::package::opc_capture_limits(limits)?;
    let root = PackURI::new("/").map_err(|error| Error::InvalidUri(error.to_string()))?;
    let current_root = package.source_relationships_with_limits(&root, opc_limits)?;
    if current_root != source.root_relationships {
        package.try_replace_relationships_with_limits(
            &current_root,
            &source.root_relationships,
            opc_limits,
        )?;
    }
    let workbook = PackURI::new(source.workbook_name())
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
    let current_workbook = package.source_relationships_with_limits(&workbook, opc_limits)?;
    if current_workbook != source.workbook_relationships {
        package.try_replace_relationships_with_limits(
            &current_workbook,
            &source.workbook_relationships,
            opc_limits,
        )?;
    }
    if let Some(model) = source.model_part() {
        let owner =
            PackURI::new(&model.name).map_err(|error| Error::InvalidUri(error.to_string()))?;
        let current_owner = package.source_relationships_with_limits(&owner, opc_limits)?;
        if current_owner != model.relationships {
            package.try_replace_relationships_with_limits(
                &current_owner,
                &model.relationships,
                opc_limits,
            )?;
        }
    }
    let current_content_types = package.source_content_types_with_limits(opc_limits)?;
    if current_content_types != source.content_types {
        package.try_replace_content_types_with_limits(
            current_content_types.bytes(),
            &source.content_types,
            opc_limits,
        )?;
    }
    Ok(())
}

/// Detached result containing a planned snapshot and reversible patch.
#[derive(Clone, Debug)]
pub struct Commit {
    patch: Patch,
    changed: bool,
}

impl Commit {
    pub(crate) fn new(patch: Patch, changed: bool) -> Self {
        Self { patch, changed }
    }

    /// Whether the transaction changes owned bytes or topology.
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Planned resulting snapshot.
    pub fn snapshot(&self) -> &Snapshot {
        self.patch.after()
    }

    /// Reversible source-checked patch.
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the result into its planned snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        let snapshot = self.patch.after().clone();
        (snapshot, self.patch)
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
