//! Immutable source-bound XLSB Data Model snapshots.

use std::sync::Arc;

use litchi_opc::{OpcPackage, OwnedContentTypes, OwnedRelationships, Part};

use super::codec::{self, ReadLimits};
use super::model::{Definition, Model, ModelPart};
use super::package::{inspect_model_part, validate_definition_connections};
use crate::package::error::{Error, Result};

/// Immutable typed metadata plus a lazy opaque model-part view.
#[derive(Clone, Debug)]
pub struct Snapshot {
    model: Option<Model>,
    connection_names: Option<Vec<String>>,
    limits: ReadLimits,
    source: SourceState,
    package: Arc<OpcPackage>,
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.model == other.model
            && self.connection_names == other.connection_names
            && self.limits == other.limits
            && self.source == other.source
    }
}

impl Eq for Snapshot {}

impl Snapshot {
    /// Read a Data Model snapshot with conservative finite limits.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::read_with_limits(package, ReadLimits::DEFAULT)
    }

    /// Read a Data Model snapshot with explicit finite limits.
    pub fn read_with_limits(package: &OpcPackage, limits: ReadLimits) -> Result<Self> {
        let limits = limits.validate()?;
        super::package::preflight_metadata(package, limits)?;
        crate::package::connections::package::preflight(
            package,
            limits.max_connection_bytes,
            limits.max_connections,
        )?;
        let workbook = package.main_document_part()?;
        if workbook.blob().len() > limits.max_part_bytes {
            return Err(Error::LimitExceeded {
                resource: "Data Model Workbook bytes",
                actual: workbook.blob().len(),
                maximum: limits.max_part_bytes,
            });
        }
        let parsed = codec::parse_workbook(workbook.blob(), limits)?;
        let part = inspect_model_part(package)?;
        let connections = crate::package::connections::package::load(package)?;
        if connections
            .as_ref()
            .is_some_and(|value| value.connections.len() > limits.max_connections)
        {
            let actual = connections
                .as_ref()
                .map_or(0, |value| value.connections.len());
            return Err(Error::LimitExceeded {
                resource: "Data Model External Data Connections",
                actual,
                maximum: limits.max_connections,
            });
        }
        let connection_source = crate::package::connections::package::capture_source(package)?;
        let model = match (parsed.definition, part) {
            (None, None) => None,
            (Some(_), None) => {
                return Err(invalid(
                    "BrtBeginDataModel exists without the relationship-free Data Model part",
                ));
            },
            (None, Some(_)) => {
                return Err(invalid(
                    "Data Model part exists without a BrtBeginDataModel workbook block",
                ));
            },
            (Some(definition), Some(part)) => Some(Model { definition, part }),
        };
        if let Some(model) = model.as_ref() {
            validate_definition_connections(&model.definition, connections.as_ref())?;
        }
        let connection_names = connections.map(|value| {
            value
                .connections
                .into_iter()
                .map(|connection| connection.name)
                .collect()
        });
        let source =
            SourceState::capture(package, workbook, model.as_ref(), connection_source, limits)?;
        Ok(Self {
            model,
            connection_names,
            limits,
            source,
            package: Arc::new(package.clone()),
        })
    }

    /// Borrow the complete model, or `None` when this workbook has no Data Model.
    pub fn model(&self) -> Option<&Model> {
        self.model.as_ref()
    }

    /// Borrow typed workbook metadata, or `None` when this workbook has no Data Model.
    pub fn definition(&self) -> Option<&Definition> {
        self.model.as_ref().map(|model| &model.definition)
    }

    /// Borrow the relationship-free opaque model payload, when present.
    pub fn part(&self) -> Option<&ModelPart> {
        self.model.as_ref().map(|model| &model.part)
    }

    /// Borrow the workbook connection names proven by the connections owner.
    ///
    /// A `None` result means that the workbook has no `/xl/connections.bin`
    /// owner. An empty slice means that the owner exists and declares no
    /// connections.
    pub fn connection_names(&self) -> Option<&[String]> {
        self.connection_names.as_deref()
    }

    /// Prove one workbook time-grouping record against the source XLDM part.
    ///
    /// The returned bindings are qualified by the XLDM table XML name and
    /// raw-column identity.  This is an explicit proof operation: reading a
    /// snapshot does not copy or decode the opaque model part, and a caller
    /// receives an error for non-version-140, incomplete, or ambiguous inner
    /// closure data.
    pub fn prove_time_grouping(
        &self,
        grouping: &super::model::TimeGrouping,
    ) -> Result<litchi_xldm::Xldm140TimeGroupingBinding> {
        let part = self
            .part()
            .ok_or_else(|| invalid("cannot prove a time grouping without a Data Model part"))?;
        super::proof::prove_time_grouping(part.bytes(), grouping)
    }

    /// Whether this workbook carries a complete Data Model pair.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.model.is_some()
    }

    /// Begin a detached, source-bound Data Model transaction.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Exact finite limits used to inspect and edit this snapshot.
    #[must_use]
    pub const fn limits(&self) -> ReadLimits {
        self.limits
    }

    pub(crate) fn source(&self) -> &SourceState {
        &self.source
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    pub(crate) fn package(&self) -> &Arc<OpcPackage> {
        &self.package
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceState {
    workbook_name: String,
    workbook: Arc<Vec<u8>>,
    model_part: Option<SourcePart>,
    connections: crate::package::connections::package::SourceImage,
    pub(crate) root_relationships: OwnedRelationships,
    pub(crate) workbook_relationships: OwnedRelationships,
    pub(crate) content_types: OwnedContentTypes,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourcePart {
    pub(crate) name: String,
    pub(crate) content_type: String,
    pub(crate) bytes: Arc<Vec<u8>>,
    pub(crate) relationships: OwnedRelationships,
}

impl SourceState {
    pub(crate) fn capture(
        package: &OpcPackage,
        workbook: &dyn Part,
        model: Option<&Model>,
        connections: crate::package::connections::package::SourceImage,
        limits: ReadLimits,
    ) -> Result<Self> {
        let opc_limits = super::package::opc_capture_limits(limits)?;
        let root_name =
            litchi_opc::PackURI::new("/").map_err(|error| Error::InvalidUri(error.to_string()))?;
        let workbook_name = workbook.partname();
        let root_relationships =
            package.source_relationships_with_limits(&root_name, opc_limits)?;
        let workbook_relationships =
            package.source_relationships_with_limits(workbook_name, opc_limits)?;
        let content_types = package.source_content_types_with_limits(opc_limits)?;
        let model_part = model
            .map(|model| {
                let part_name = litchi_opc::PackURI::new(&model.part.part_name)
                    .map_err(|error| Error::InvalidUri(error.to_string()))?;
                Ok::<SourcePart, Error>(SourcePart {
                    name: model.part.part_name.clone(),
                    content_type: model.part.content_type.clone(),
                    bytes: Arc::clone(&model.part.bytes),
                    relationships: package
                        .source_relationships_with_limits(&part_name, opc_limits)?,
                })
            })
            .transpose()?;
        Ok(Self {
            workbook_name: workbook.partname().to_string(),
            workbook: workbook.blob_arc(),
            model_part,
            connections,
            root_relationships,
            workbook_relationships,
            content_types,
        })
    }

    pub(crate) fn workbook_name(&self) -> &str {
        &self.workbook_name
    }

    pub(crate) fn model_part(&self) -> Option<&SourcePart> {
        self.model_part.as_ref()
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
