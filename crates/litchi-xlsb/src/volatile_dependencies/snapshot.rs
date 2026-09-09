//! Immutable source-bound Volatile Dependencies snapshots.

use std::sync::Arc;

use litchi_opc::{OpcPackage, OwnedContentTypes, OwnedRelationships, Part};

use super::codec;
use super::model::{Dependencies, ReadLimits};
use super::package::{self, Graph};
use crate::package::error::{Error, Result};
use crate::raw::kind;

/// An immutable typed view over one Workbook-owned Volatile Dependencies part.
#[derive(Clone, Debug)]
pub struct Snapshot {
    dependencies: Option<Arc<Dependencies>>,
    sheet_count: usize,
    limits: ReadLimits,
    source: SourceState,
    package: Arc<OpcPackage>,
    rewrite_safe: bool,
}

impl Snapshot {
    /// Read a snapshot with explicit finite limits.
    pub(crate) fn read_with_limits(package: &OpcPackage, limits: ReadLimits) -> Result<Self> {
        let limits = limits.validate()?;
        let workbook = package.main_document_part()?;
        let workbook_name = workbook.partname().clone();
        let sheet_count = workbook_sheet_count(workbook, limits)?;
        let graph = package::discover_graph(package, &workbook_name, limits)?;
        let (dependencies, rewrite_safe) = if let Some(graph) = graph.as_ref() {
            let part = package.get_part(&graph.part_name)?;
            if part.blob().len() > limits.max_part_bytes {
                return Err(Error::LimitExceeded {
                    resource: "Volatile Dependencies part bytes",
                    actual: part.blob().len(),
                    maximum: limits.max_part_bytes,
                });
            }
            let parsed = codec::read(part.blob(), limits, sheet_count)?;
            (Some(Arc::new(parsed.dependencies)), parsed.rewrite_safe)
        } else {
            (None, true)
        };
        let source = SourceState::capture(package, &workbook_name, graph.as_ref(), limits)?;
        Ok(Self {
            dependencies,
            sheet_count,
            limits,
            source,
            package: Arc::new(package.clone()),
            rewrite_safe,
        })
    }

    /// Borrow the typed hierarchy, or `None` when the optional part is absent.
    #[must_use]
    pub fn dependencies(&self) -> Option<&Dependencies> {
        self.dependencies.as_deref()
    }

    /// Whether the Workbook owns a Volatile Dependencies part.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.dependencies.is_some()
    }

    /// Whether a typed replacement can preserve every source record.
    #[must_use]
    pub const fn can_edit(&self) -> bool {
        self.rewrite_safe
    }

    /// Start a detached source-bound transaction. Missing owners can be
    /// created through `replace` on the returned transaction.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Exact finite limits retained by this snapshot.
    #[must_use]
    pub const fn limits(&self) -> ReadLimits {
        self.limits
    }

    pub(crate) const fn sheet_count(&self) -> usize {
        self.sheet_count
    }

    pub(crate) fn source(&self) -> &SourceState {
        &self.source
    }

    pub(crate) fn package(&self) -> &Arc<OpcPackage> {
        &self.package
    }

    pub(crate) fn dependencies_arc(&self) -> Option<Arc<Dependencies>> {
        self.dependencies.as_ref().map(Arc::clone)
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    pub(crate) fn same_state(&self, other: &Self) -> bool {
        self.dependencies == other.dependencies && self.same_source(other)
    }
}

/// Count the `BrtBundleSh` records directly contained by the Workbook's
/// `BrtBeginBundleShs` collection. `BrtVolRef.ish` is defined as this
/// zero-based ordinal, so semantic Volatile inspection cannot accept an
/// arbitrary sheet number.
fn workbook_sheet_count(workbook: &dyn Part, limits: ReadLimits) -> Result<usize> {
    let data = workbook.blob();
    if data.len() > limits.max_part_bytes {
        return Err(Error::LimitExceeded {
            resource: "Workbook bytes while binding volatile sheet ordinals",
            actual: data.len(),
            maximum: limits.max_part_bytes,
        });
    }
    let raw_limits = crate::raw::Limits::new(limits.max_part_bytes, limits.max_string_units);
    let mut records = crate::raw::Records::try_with_limits(data, raw_limits)?;
    let mut record_count = 0usize;
    let mut sheet_count = 0usize;
    let mut collection_seen = false;
    let mut collection_open = false;
    while let Some(record) = records.next() {
        let record = record?;
        record_count = record_count.checked_add(1).ok_or(Error::CapacityOverflow {
            resource: "Workbook records while binding volatile sheet ordinals",
        })?;
        if record_count > limits.max_records {
            return Err(Error::LimitExceeded {
                resource: "Workbook records while binding volatile sheet ordinals",
                actual: record_count,
                maximum: limits.max_records,
            });
        }
        match record.kind() {
            kind::BEGIN_BUNDLE_SHS => {
                if !record.payload().is_empty() {
                    return Err(invalid("BrtBeginBundleShs has a non-empty payload"));
                }
                if collection_seen || collection_open {
                    return Err(invalid("duplicate or nested BrtBeginBundleShs"));
                }
                collection_seen = true;
                collection_open = true;
            },
            kind::BUNDLE_SH => {
                if !collection_open {
                    return Err(invalid("BrtBundleSh is outside BrtBeginBundleShs"));
                }
                sheet_count = sheet_count.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "Workbook BrtBundleSh records",
                })?;
            },
            kind::END_BUNDLE_SHS => {
                if !record.payload().is_empty() {
                    return Err(invalid("BrtEndBundleShs has a non-empty payload"));
                }
                if !collection_open {
                    return Err(invalid("BrtEndBundleShs has no matching begin"));
                }
                collection_open = false;
            },
            _ if collection_open => {
                return Err(invalid(
                    "Workbook BrtBundleShs collection contains a non-direct record",
                ));
            },
            _ => {},
        }
    }
    if collection_open {
        return Err(Error::UnexpectedEndOfStream(
            "Workbook BrtBeginBundleShs collection".to_string(),
        ));
    }
    Ok(sheet_count)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.dependencies == other.dependencies && self.source == other.source
    }
}

impl Eq for Snapshot {}

/// Exact source closure owned by this package owner.
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
        let relationship_count = workbook
            .rels()
            .len()
            .checked_add(package.rels().len())
            .ok_or(Error::CapacityOverflow {
                resource: "Volatile Dependencies relationships",
            })?;
        if relationship_count > limits.max_relationships {
            return Err(Error::LimitExceeded {
                resource: "Volatile Dependencies relationships",
                actual: relationship_count,
                maximum: limits.max_relationships,
            });
        }
        let workbook_relationships = package.source_relationships(workbook_name)?;
        let root_name =
            litchi_opc::PackURI::new("/").map_err(|error| Error::InvalidUri(error.to_string()))?;
        let root_relationships = package.source_relationships(&root_name)?;
        let content_types = package.source_content_types()?;
        if workbook_relationships.bytes().len() > limits.max_part_bytes
            || root_relationships.bytes().len() > limits.max_part_bytes
            || content_types.bytes().len() > limits.max_part_bytes
        {
            return Err(Error::LimitExceeded {
                resource: "Volatile Dependencies lexical source bytes",
                actual: workbook_relationships
                    .bytes()
                    .len()
                    .max(root_relationships.bytes().len())
                    .max(content_types.bytes().len()),
                maximum: limits.max_part_bytes,
            });
        }
        let workbook = SourcePart {
            name: workbook_name.as_str().to_string(),
            content_type: workbook.content_type().to_string(),
            bytes: workbook.blob_arc(),
            relationships: workbook_relationships.clone(),
        };
        let owner = graph
            .map(|graph| package.get_part(&graph.part_name))
            .transpose()?
            .map(|part| {
                let relationships = package.source_relationships(part.partname())?;
                Ok::<SourcePart, Error>(SourcePart {
                    name: part.partname().as_str().to_string(),
                    content_type: part.content_type().to_string(),
                    bytes: part.blob_arc(),
                    relationships,
                })
            })
            .transpose()?;
        Ok(Self {
            graph: graph.cloned(),
            workbook,
            root_relationships,
            workbook_relationships,
            content_types,
            owner,
        })
    }
}
