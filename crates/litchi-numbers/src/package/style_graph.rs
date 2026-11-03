//! The bridge between a Numbers package and W1-E's shared style-variation
//! graph engine (`litchi_iwa_style::graph`, record 0375a), for the chart
//! series families of record 0378: the package's parsed components as the
//! engine's [`Components`], and a finished workspace's [`Changes`]
//! republished through exact-source reassembly and a strict reopen.

use std::mem::size_of;
use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, EntryInsertion};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_style::graph::{Changes, Components, Error as GraphError};

use super::Package;
use super::structure::{self, Fault, FaultResult, Published};

/// Graph faults retained by the style-family adapters without widening the
/// shared structural fault taxonomy.
#[derive(Debug, Clone, Copy)]
pub(super) enum StyleGraphFault {
    Shared(Fault),
    Limit {
        kind: litchi_iwa_style::graph::GraphLimitKind,
        observed: u64,
        maximum: u64,
    },
}

/// The package's parsed components, in package order.
pub(super) struct LeafComponents<'p> {
    components: Vec<(&'p str, &'p Archive)>,
}

impl<'p> LeafComponents<'p> {
    pub(super) fn view_requirements(package: &Package) -> FaultResult<(usize, usize)> {
        structure::exact_source(package)?;
        let count = package.state.components.catalog().len();
        let bytes = count
            .checked_mul(size_of::<(&str, &Archive)>())
            .and_then(|value| value.checked_add(size_of::<LeafComponents<'_>>()));
        let bytes = bytes.ok_or(Fault::Invalid)?;
        Ok((usize::from(count != 0) + 1, bytes))
    }

    pub(super) fn new(package: &'p Package) -> FaultResult<Self> {
        structure::exact_source(package)?;
        let mut components = Vec::new();
        structure::reserve(&mut components, package.state.components.catalog().len())?;
        for component in package.state.components.catalog().iter() {
            components.push((component.name(), component.archive()));
        }
        Ok(Self { components })
    }
}

impl Components for LeafComponents<'_> {
    fn count(&self) -> usize {
        self.components.len()
    }

    fn component(&self, index: usize) -> Option<(&str, &Archive)> {
        self.components
            .get(index)
            .map(|(name, archive)| (*name, *archive))
    }
}

/// The fault a graph-engine failure maps to.
pub(super) fn fault(error: &GraphError) -> StyleGraphFault {
    match error {
        GraphError::Archive(error) => StyleGraphFault::Shared(structure::map_core(error)),
        GraphError::Wire(error) => StyleGraphFault::Shared(structure::map_wire(error)),
        GraphError::IdentifierOverflow => StyleGraphFault::Shared(Fault::Limit {
            kind: structure::FaultLimit::Identifiers,
            observed: u64::MAX,
            maximum: u64::MAX,
        }),
        GraphError::Limit {
            kind: litchi_iwa_style::graph::GraphLimitKind::ScratchBytes,
            observed,
            maximum,
        } => StyleGraphFault::Limit {
            kind: litchi_iwa_style::graph::GraphLimitKind::ScratchBytes,
            observed: *observed as u64,
            maximum: *maximum as u64,
        },
        GraphError::Limit {
            kind: litchi_iwa_style::graph::GraphLimitKind::AllocationEvents,
            observed,
            maximum,
        } => StyleGraphFault::Limit {
            kind: litchi_iwa_style::graph::GraphLimitKind::AllocationEvents,
            observed: *observed as u64,
            maximum: *maximum as u64,
        },
        _ => StyleGraphFault::Shared(Fault::Invalid),
    }
}

/// Republish `changes` over `package` once and reopen the candidate.
/// Publish `changes` and the new non-IWA members `data` (embedded media
/// a data-owner transition registered).
pub(super) fn publish(
    package: &Package,
    changes: Changes,
    data: &[(&str, &[u8])],
) -> FaultResult<Published> {
    let catalog = structure::exact_source(package)?;
    let limits = catalog.limits();
    let archive_limits = limits
        .effective_archive_limits()
        .map_err(|error| structure::map_archive(&error))?;
    let snappy = limits
        .snappy_limits()
        .map_err(|error| structure::map_archive(&error))?;
    let compress = |archive: &Archive| -> FaultResult<Vec<u8>> {
        let bytes = archive
            .to_bytes(archive_limits)
            .map_err(|error| structure::map_core(&error))?;
        SnappyStream::compress(&bytes, snappy).map_err(|error| structure::map_core(&error))
    };
    let mut edited = Vec::new();
    for (name, archive) in &changes.edits {
        structure::reserve(&mut edited, 1)?;
        edited.push((name.as_str(), compress(archive)?));
    }
    let names: Vec<&str> = catalog.package().iter().map(|entry| entry.name()).collect();
    let mut inserted = Vec::new();
    for insertion in &changes.insertions {
        let position = match &insertion.before {
            Some(before) => Some(
                names
                    .iter()
                    .position(|name| name == before)
                    .ok_or(Fault::Invalid)?,
            ),
            None => None,
        };
        structure::reserve(&mut inserted, 1)?;
        inserted.push((
            insertion.name.as_str(),
            compress(&insertion.archive)?,
            position,
        ));
    }
    let edits: Vec<EntryEdit<'_>> = edited
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect();
    let insertions: Vec<EntryInsertion<'_>> = inserted
        .iter()
        .map(|(name, data, position)| match position {
            Some(position) => EntryInsertion::at(*position, name, data),
            None => EntryInsertion::new(name, data),
        })
        .chain(
            data.iter()
                .map(|(name, bytes)| EntryInsertion::new(name, bytes)),
        )
        .collect();
    let deletions: Vec<&str> = changes.deletions.iter().map(String::as_str).collect();
    let output = catalog
        .package()
        .reassemble_with_changes_to_bytes(&insertions, &edits, &deletions, limits)
        .map_err(|error| structure::map_archive(&error))?;
    let candidate = package
        .reopen_candidate(Arc::from(output))
        .map_err(|_error| Fault::Verification)?;
    Ok(Published {
        package: candidate,
        touched_components: edits.len() + insertions.len() + deletions.len(),
        deleted_previews: 0,
    })
}
