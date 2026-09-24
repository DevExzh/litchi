//! Immutable source-bound Calculation Chain snapshots.

use std::sync::Arc;

use litchi_opc::{OpcPackage, OwnedContentTypes, OwnedRelationships};

use super::model::{ChainPart, ReadLimits};
use super::package::{self, Graph};
use crate::package::error::{Error, Result};

/// An immutable opaque view over one Workbook-owned Calculation Chain part.
#[derive(Clone, Debug)]
pub struct Snapshot {
    part: Option<Arc<ChainPart>>,
    limits: ReadLimits,
    source: SourceState,
    package: Arc<OpcPackage>,
}

impl Snapshot {
    /// Read a snapshot with explicit finite limits.
    pub(crate) fn read_with_limits(package: &OpcPackage, limits: ReadLimits) -> Result<Self> {
        Self::read_with_probe(package, limits, None)
    }

    /// Read a snapshot while optionally proving that a future owner target has
    /// no hidden inbound relationship when the current owner is absent.
    pub(crate) fn read_with_probe(
        package: &OpcPackage,
        limits: ReadLimits,
        probe_target: Option<&litchi_opc::PackURI>,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        // `discover_graph_with_probe` performs the borrowed package metadata
        // preflight before it returns any owned graph state.  Do not resolve
        // or clone the Workbook URI before that guard has run.
        let graph = package::discover_graph_with_probe(package, limits, probe_target)?;
        let workbook = package.main_document_part()?;
        let workbook_name = workbook.partname().clone();
        if workbook.blob().len() > limits.max_part_bytes {
            return Err(Error::LimitExceeded {
                resource: "Calculation Chain Workbook bytes",
                actual: workbook.blob().len(),
                maximum: limits.max_part_bytes,
            });
        }
        let part = graph
            .as_ref()
            .map(|graph| package.get_part(&graph.part_name))
            .transpose()?
            .map(|part| {
                let bytes = part.blob_arc();
                if bytes.len() > limits.max_part_bytes {
                    return Err(Error::LimitExceeded {
                        resource: "Calculation Chain part bytes",
                        actual: bytes.len(),
                        maximum: limits.max_part_bytes,
                    });
                }
                Ok(ChainPart {
                    part_name: part.partname().as_str().to_string(),
                    content_type: part.content_type().to_string(),
                    bytes,
                    max_records: limits.max_records,
                })
            })
            .transpose()?
            .map(Arc::new);
        let source = SourceState::capture(package, &workbook_name, graph.as_ref(), limits)?;
        Ok(Self {
            part,
            limits,
            source,
            package: Arc::new(package.clone()),
        })
    }

    /// Borrow the opaque physical part, or `None` when the Workbook does not
    /// own a Calculation Chain part.
    #[must_use]
    pub fn part(&self) -> Option<&ChainPart> {
        self.part.as_deref()
    }

    /// Whether the Workbook owns a Calculation Chain part.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.part.is_some()
    }

    /// Start a detached source-bound transaction. The only supported change
    /// is removal of the optional cache owner.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Exact finite limits retained by this snapshot.
    #[must_use]
    pub const fn limits(&self) -> ReadLimits {
        self.limits
    }

    pub(crate) fn source(&self) -> &SourceState {
        &self.source
    }

    pub(crate) fn package(&self) -> &Arc<OpcPackage> {
        &self.package
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    pub(crate) fn same_state(&self, other: &Self) -> bool {
        self.part == other.part && self.same_source(other)
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.same_state(other)
    }
}

impl Eq for Snapshot {}

/// Exact source closure used for stale checks and lexical restoration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceState {
    pub(crate) graph: Option<Graph>,
    pub(crate) workbook: SourcePart,
    pub(crate) root_relationships: OwnedRelationships,
    pub(crate) workbook_relationships: OwnedRelationships,
    pub(crate) content_types: OwnedContentTypes,
    pub(crate) owner: Option<SourcePart>,
}

/// One source part and its lexical relationship state.
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
        workbook_name: &litchi_opc::PackURI,
        graph: Option<&Graph>,
        limits: ReadLimits,
    ) -> Result<Self> {
        let workbook = package.get_part(workbook_name)?;
        let root_name =
            litchi_opc::PackURI::new("/").map_err(|error| Error::InvalidUri(error.to_string()))?;
        let opc_limits = opc_capture_limits(limits)?;
        let root_relationships =
            package.source_relationships_with_limits(&root_name, opc_limits)?;
        let workbook_relationships =
            package.source_relationships_with_limits(workbook_name, opc_limits)?;
        let content_types = package.source_content_types_with_limits(opc_limits)?;
        let owner = graph
            .map(|graph| package.get_part(&graph.part_name))
            .transpose()?
            .map(|part| {
                let relationships =
                    package.source_relationships_with_limits(part.partname(), opc_limits)?;
                Ok::<SourcePart, Error>(SourcePart {
                    name: part.partname().as_str().to_string(),
                    content_type: part.content_type().to_string(),
                    bytes: part.blob_arc(),
                    relationships,
                })
            })
            .transpose()?;

        let mut maximum = workbook.blob().len();
        maximum = maximum.max(root_relationships.bytes().len());
        maximum = maximum.max(workbook_relationships.bytes().len());
        maximum = maximum.max(content_types.bytes().len());
        if let Some(owner) = owner.as_ref() {
            maximum = maximum.max(owner.bytes.len());
            maximum = maximum.max(owner.relationships.bytes().len());
        }
        if maximum > limits.max_part_bytes {
            return Err(Error::LimitExceeded {
                resource: "Calculation Chain source closure bytes",
                actual: maximum,
                maximum: limits.max_part_bytes,
            });
        }

        Ok(Self {
            graph: graph.cloned(),
            workbook: SourcePart {
                name: workbook_name.as_str().to_string(),
                content_type: workbook.content_type().to_string(),
                bytes: workbook.blob_arc(),
                relationships: workbook_relationships.clone(),
            },
            root_relationships,
            workbook_relationships,
            content_types,
            owner,
        })
    }
}

pub(crate) fn opc_capture_limits(limits: ReadLimits) -> Result<litchi_opc::ReadLimits> {
    if limits.max_part_bytes == 0 {
        return Err(Error::LimitExceeded {
            resource: "Calculation Chain source closure bytes",
            actual: 1,
            maximum: limits.max_part_bytes,
        });
    }
    let maximum_relationship_xml =
        limits
            .max_part_bytes
            .checked_mul(3)
            .ok_or(Error::CapacityOverflow {
                resource: "Calculation Chain relationship source bytes",
            })?;
    let defaults = litchi_opc::ReadLimits::default();
    let maximum_relationships = limits.max_relationships.max(1);
    let maximum_parts = limits.max_graph_parts.max(1).min(defaults.max_parts());
    let maximum_mappings = limits.max_graph_parts.max(1).min(
        defaults
            .max_content_type_mappings()
            .max(defaults.max_parts()),
    );
    let builder = litchi_opc::ReadLimits::builder()
        .max_parts(maximum_parts)?
        .max_content_types_bytes(limits.max_part_bytes)?
        .max_content_type_mappings(maximum_mappings)?
        .max_relationship_parts(3)?
        .max_relationship_xml_bytes(limits.max_part_bytes)?
        .max_total_relationship_xml_bytes(maximum_relationship_xml)?
        .max_relationships_per_part(maximum_relationships)?
        .max_total_relationships(maximum_relationships)?
        .max_relationship_graph_nodes(maximum_parts)?;
    Ok(builder.build()?)
}
